import fs from 'node:fs/promises';
import path from 'node:path';
import net from 'node:net';
import {fileURLToPath} from 'node:url';
import {parseOrder, example, gasketExample} from './parser.mjs';
import {generate} from './excel.mjs';
import {connect} from './startup.mjs';
import {latestDueWeek, loadOrders, nextWeek, orderNotice, summaryText, summaryWorkbook, weekOf} from './summary.mjs';
import {archiveDeleted, deleteBody, deleteSubject, findOrder, mailBody, mailSubject, sendMail} from './mail.mjs';
import {generateSteelForms, isSteelForm, parseSteelForm, steelExample, steelNotice} from './steel.mjs';
import {gasketOrderText, gasketRowsIn, gridToOrderLines, imageToLines, imageToXlsx, linesToGrid, isGasketHeader, isGasketTable, orderTextFromTable, readGrid} from './table.mjs';
const root=path.dirname(fileURLToPath(import.meta.url));
const token=process.env.TELEGRAM_BOT_TOKEN;
delete process.env.TELEGRAM_BOT_TOKEN;
if (!token) {console.error('Run Setup.cmd first, then Start.cmd.'); process.exit(1);}
// Local lock prevents two polling instances without persisting a stale PID lock.
const lock=net.createServer();
lock.on('error',()=>{console.error('Bot already running, or local port 43871 is in use.'); process.exit(1);});
await new Promise(resolve=>lock.listen(43871,'127.0.0.1',resolve));
const cfg=JSON.parse((await fs.readFile(path.join(root,'config.json'),'utf8')).replace(/^\uFEFF/,''));
let state;
try {state=JSON.parse((await fs.readFile(path.join(root,'state.json'),'utf8')).replace(/^\uFEFF/,''));}
catch(e) {if(e.code!=='ENOENT') throw e; state={offset:0,owner:null,users:[],pairFailures:0};}
async function save() {
  await fs.writeFile(path.join(root,'state.tmp'),JSON.stringify(state,null,2));
  await fs.rename(path.join(root,'state.tmp'),path.join(root,'state.json'));
}
async function api(method,data,multipart=false) {
  const result=await fetch(`https://api.telegram.org/bot${token}/${method}`,{method:'POST',headers:multipart?{}:{'Content-Type':'application/json'},body:multipart?data:JSON.stringify(data),signal:AbortSignal.timeout(65000)});
  const payload=await result.json();
  if (!payload.ok) {const e=Error(`Telegram ${method} failed (${payload.error_code})`); e.code=payload.error_code; e.retryAfter=payload.parameters?.retry_after; throw e;}
  return payload.result;
}
const say=(chat,text)=>api('sendMessage',{chat_id:chat,text});
const MIME={'.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','.pdf':'application/pdf'};
async function sendFile(chat,filename,caption) {
  const form=new FormData();
  form.append('chat_id',String(chat));
  if(caption) form.append('caption',caption);
  form.append('document',new Blob([await fs.readFile(filename)],{type:MIME[path.extname(filename).toLowerCase()] || 'application/octet-stream'}),path.basename(filename));
  return api('sendDocument',form,true);
}
// Saves a file the user sent (photo or document) into the job folder.
async function download(fileId,dest) {
  const info=await api('getFile',{file_id:fileId});
  const res=await fetch(`https://api.telegram.org/file/bot${token}/${info.file_path}`,{signal:AbortSignal.timeout(65000)});
  if(!res.ok) throw Error(`Telegram file download failed (${res.status})`);
  await fs.writeFile(dest,Buffer.from(await res.arrayBuffer()));
}
const {me,webhook}=await connect(api);
if (webhook.url) {console.error('This bot already has a webhook. Use a new BotFather bot or remove its existing integration intentionally.');process.exit(1);}
console.log(`Online: @${me.username}. Private chats only. ${new Date().toISOString()}`);
await fs.writeFile(path.join(root,'status.json'),JSON.stringify({username:me.username,started:new Date().toISOString(),pid:process.pid}));

// Weekly order summaries go only to the administrator: a message plus an Excel list kept in summaries/.
async function sendSummary(chat,week,soFar=false) {
  const orders=await loadOrders(path.join(root,'outputs'),week);
  await say(chat,summaryText(orders,week,soFar));
  if(!orders.length) return;
  await sendFile(chat,await summaryWorkbook(orders,week,path.join(root,'summaries')));
}
let summaryRetryAt=0;
// Sends each week at 09:00 on the following Monday (Singapore time), including weeks missed while the host was off.
async function sendDueSummaries() {
  if(!state.owner || Date.now()<summaryRetryAt) return;
  const due=latestDueWeek();
  if(!state.summaryWeek) {state.summaryWeek=due; await save(); return;}
  try {
    while(state.summaryWeek<due) {
      const week=nextWeek(state.summaryWeek);
      await sendSummary(state.owner,week);
      state.summaryWeek=week; await save();
    }
  } catch(e) {
    console.error(`Weekly summary failed (${e.code || 'network/local error'}). ${new Date().toISOString()}`);
    summaryRetryAt=Date.now()+600000;
  }
}

// Steel-work Word forms: the intake from the steel-forms handoff, filled with its own updater and approved
// templates (folder and interpreter in steel.json). Both Word files and Word-rendered PDF previews go back
// to the sender; a colleague's forms are also copied to the administrator.
async function handleSteel(m,text,sender,uid,chat) {
  let cfg;
  try {cfg=JSON.parse((await fs.readFile(path.join(root,'steel.json'),'utf8')).replace(/^﻿/,''));} catch(e) {if(e.code!=='ENOENT') throw e;}
  if(!cfg?.root) {await say(chat,'Steel work forms are not set up on the host computer. Please contact the administrator.');return;}
  let sub;
  try {sub=parseSteelForm(text);}
  catch(e) {await say(chat,e.message+'\n\nPlease correct and resend the complete form. Send /help for the layout.');return;}
  const id=`${chat}_${m.message_id}`;
  const jobdir=path.join(root,'steel',id);
  await fs.mkdir(jobdir,{recursive:true});
  const jobfile=path.join(jobdir,'delivery.json');
  let prior;
  try {prior=JSON.parse(await fs.readFile(jobfile,'utf8'));} catch(e) {if(e.code!=='ENOENT') throw e;}
  if(prior?.status==='sent') return;
  if(prior?.status) {await say(chat,'The previous delivery of these forms could not be confirmed. Please check whether you received them. If not, resend the form.');return;}
  await fs.writeFile(path.join(jobdir,'request.json'),JSON.stringify({telegramUserId:uid,receivedAt:new Date(m.date*1000).toISOString(),submission:sub,source:text},null,2));
  let result;
  try {result=await generateSteelForms(sub,cfg);}
  catch(e) {console.error(`Steel forms failed for ${id}: ${String(e.message).split(/\r?\n/)[0]} ${new Date().toISOString()}`);await say(chat,'The steel work forms could not be generated. Please contact the administrator. Your submission has been saved on the host computer.');return;}
  await fs.writeFile(path.join(jobdir,'result.json'),JSON.stringify({outDir:result.outDir,jsonFile:result.jsonFile,files:result.files,verify:result.verify,layout:result.layout},null,2));
  if(!result.verify.ok) {console.error(`Steel forms for ${id} failed the field check: ${result.verify.problems.join(' | ')}`);await say(chat,'The steel work forms were generated but did not pass the field check, so they were not sent. Please contact the administrator.');return;}
  const caption=steelNotice(sub,result);
  await fs.writeFile(jobfile,JSON.stringify({status:'sending'}));
  let delivery;
  try {
    await sendFile(chat,result.files.written,caption);
    await sendFile(chat,result.files.request);
    for (const pdf of Object.values(result.layout?.pdfs || {})) if(pdf) await sendFile(chat,pdf,'Preview as rendered by Word on the host computer');
    delivery={status:'sent'};
    await fs.writeFile(jobfile,JSON.stringify(delivery));
    console.log(`Steel forms delivered ${id} (${sub.ref})`);
  } catch {
    delivery={status:'uncertain'};
    await fs.writeFile(jobfile,JSON.stringify(delivery));
    await say(chat,'The forms were generated on the host computer, but Telegram delivery could not be confirmed. Please check whether you received them. If not, resend the form.');
  }
  if(!state.owner || uid===state.owner) return;
  const who=m.from.username ? `${sender} (@${m.from.username})` : sender;
  try {await sendFile(state.owner,result.files.written,`New steel work forms from ${who}\n${caption}`.slice(0,1024)); await sendFile(state.owner,result.files.request); delivery.adminCopy='sent';}
  catch(e) {console.error(`Administrator copy failed for steel ${id} (${e.code || 'network/local error'}). ${new Date().toISOString()}`); delivery.adminCopy='failed';}
  await fs.writeFile(jobfile,JSON.stringify(delivery));
}

async function handle(m) {
  if(m.chat?.type!=='private' || !m.from || m.from.is_bot) return;
  const text=(m.text || '').trim(), uid=m.from.id, chat=m.chat.id;
  if (/^\/id(?:@\w+)?$/.test(text)) {await say(chat,`Your Telegram ID: ${uid}\nSend this ID to the administrator. They can grant access using /allow ${uid}.`); return;}
  if (!state.owner && /^\/pair\s+/.test(text)) {
    if(state.pairFailures>=10) {await say(chat,'Pairing is locked. Please ask the administrator to run Setup.cmd again.');return;}
    if(text.slice(6).trim() !== cfg.pairCode) {state.pairFailures++;await save();await say(chat,'Incorrect pairing code. Please use the code shown in the latest Setup window.');return;}
    state.owner=uid; state.users=[uid]; await save();
    await say(chat,'Administrator pairing completed. Send /help for the order format. Colleagues can send /id and share their ID with you. Use /allow ID to grant them access.'); return;
  }
  if(!state.users.includes(uid)) {await say(chat,'Welcome to the TWRP C3D1/C3D2 Bot. Please contact the administrator for access. Send /id to get your Telegram ID. For initial administrator setup, send /pair followed by the pairing code.');return;}
  const allow=text.match(/^\/(allow|remove)\s+(\d+)$/);
  if(allow && uid===state.owner) {
    const id=Number(allow[2]);
    if(id===state.owner) {await say(chat,'Administrator access remains enabled.');return;}
    if(allow[1]==='allow') state.users=[...new Set([...state.users,id])]; else state.users=state.users.filter(x=>x!==id);
    await save(); await say(chat,`${allow[1]} ${id}: OK`);return;
  }
  if(/^\/summary(?:@\w+)?$/.test(text) && uid===state.owner) {
    try {await sendSummary(chat,weekOf(Date.now()),true);}
    catch(e) {console.error(`Summary request failed (${e.code || 'network/local error'}). ${new Date().toISOString()}`); await say(chat,'The order summary could not be sent. Please try again later.');}
    return;
  }
  // /sendemail emails the latest order file in this chat, or the file replied to, to the administrator's mailbox.
  if(/^\/sendemail(?:@\w+)?$/i.test(text)) {
    let mail;
    try {mail=JSON.parse((await fs.readFile(path.join(root,'email.json'),'utf8')).replace(/^﻿/,''));} catch(e) {if(e.code!=='ENOENT') throw e;}
    if(!mail?.to) {await say(chat,'Email is not set up on the host computer. Please contact the administrator.');return;}
    const sender=[m.from.first_name,m.from.last_name].filter(Boolean).join(' ');
    const found=await findOrder(path.join(root,'outputs'),{chat,owner:state.owner,replyTo:m.reply_to_message?.message_id});
    if(!found) {await say(chat,m.reply_to_message?'Please reply to an Excel file sent by this bot, or send /sendemail on its own to email your latest order.':'No generated order found in this chat yet. Send an order first, then /sendemail.');return;}
    const subject=mailSubject(found.order);
    try {
      const result=await sendMail({to:mail.to,subject,body:mailBody(found.order,sender,found.file),attachment:found.file},found.dir);
      await fs.writeFile(path.join(found.dir,'email.json'),JSON.stringify({...result,to:mail.to,subject,at:new Date().toISOString(),telegramUserId:uid}));
      await say(chat,`${result.status==='sent'?'Emailed':'Handed to Outlook for sending'} to ${mail.to}:\n${subject}`);
      console.log(`Order emailed ${path.basename(found.dir)} (${result.status})`);
    } catch(e) {
      console.error(`Email failed for ${path.basename(found.dir)}: ${e.message.split(/\r?\n/)[0]} ${new Date().toISOString()}`);
      await say(chat,'The email could not be sent. Please try again later or contact the administrator.');
    }
    return;
  }
  // /delete removes the latest order in this chat (or the order of the file replied to) from the records, because its
  // file is wrong, and tells the administrator by email with the deleted file attached. The order folder moves to deleted/.
  if(/^\/delete(?:order)?(?:@\w+)?$/i.test(text)) {
    const found=await findOrder(path.join(root,'outputs'),{chat,owner:state.owner,replyTo:m.reply_to_message?.message_id});
    if(!found) {await say(chat,m.reply_to_message?'Please reply to an Excel file sent by this bot, or send /delete on its own to delete your latest order.':'No generated order found in this chat to delete.');return;}
    const sender=[m.from.first_name,m.from.last_name].filter(Boolean).join(' ');
    const subject=mailSubject(found.order), name=path.basename(found.dir), when=new Date();
    let emailedAt;
    try {emailedAt=JSON.parse(await fs.readFile(path.join(found.dir,'email.json'),'utf8')).at;} catch {}
    let dir;
    try {dir=await archiveDeleted(found.dir,path.join(root,'deleted'),{at:when.toISOString(),telegramUserId:uid,sender,chat,emailedAt});}
    catch(e) {console.error(`Delete failed for ${name}: ${e.code||'local error'} ${when.toISOString()}`);await say(chat,'The order could not be deleted. Please contact the administrator.');return;}
    console.log(`Order deleted ${name}`);
    let mail;
    try {mail=JSON.parse((await fs.readFile(path.join(root,'email.json'),'utf8')).replace(/^﻿/,''));} catch(e) {if(e.code!=='ENOENT') throw e;}
    let notified=false;
    if(mail?.to) {
      try {
        const file=path.join(dir,path.basename(found.file));
        const result=await sendMail({to:mail.to,subject:deleteSubject(found.order),body:deleteBody(found.order,sender,file,{when,emailedAt}),attachment:file},dir);
        await fs.writeFile(path.join(dir,'deleted.json'),JSON.stringify({at:when.toISOString(),telegramUserId:uid,sender,chat,emailedAt,notice:{...result,to:mail.to}}));
        notified=true;
      } catch(e) {console.error(`Delete notice email failed for ${name}: ${e.message.split(/\r?\n/)[0]} ${new Date().toISOString()}`);}
    }
    // The administrator must always hear about a deletion: fall back to Telegram when the email could not be sent.
    if(!notified && chat!==state.owner) {try {await say(state.owner,`${sender} deleted an order (the email notice could not be sent):\n${subject}`);notified=true;} catch {}}
    await say(chat,`Deleted: ${subject}\nIt is removed from the order records.${notified?' The administrator has been told.':' The administrator could not be notified automatically, please tell them.'}\nThe file already sent in this chat stays there; please do not use it.${emailedAt?'\nThis order had already been emailed with /sendemail; the administrator was told to disregard that email.':''}`);
    return;
  }
  if (/^\/(?:start|help|template)(?:@\w+)?$/.test(text)) {
    await say(chat,`Welcome to the TWRP C3D1/C3D2 Bot. Send one complete order in a single text message to receive an Excel file.\n\nFastener example:\n${example}\n\nGasket example:\n${gasketExample}\n\nProjects: C3D1 or C3D2. Materials: SS316 (SS16 is also accepted), HDG or Gasket (EPDM). Send different projects or materials in separate messages.\nUse a separate Title for each location. Matching nuts, washers and gaskets are combined within each location; different bolt lengths remain separate.\nRequester Name defaults to your Telegram display name unless you specify it.\nFor sets: write Sets: 75 immediately after the Title, then enter the quantity per set.\nNo extra 5% is added.\nWrite the quantity after =, : or qty on each line, for example M16x2m stud rod = 70 or M16 nut qty=1000 (2m means a 2000mm rod).\nFastener orders can also be sent as an Excel file or a photo of the table: sizes down the first column (M24X210, M20x60 (Bolt only)), Bolt / Nut / Washer columns across the top, and a caption giving the project, material and Title, for example C3D1 SS16 then Title: Knife Gate Valves. A list with one item per row (HEXAGON BOLT: M12 Gr 8.8, L=45mm / NUT: M12 / WASHER: M12 X 3mm THKS) and a Total Qty column is read too, and the caption can be left out when the top row of the table gives the project and title. The bot replies with what it read so you can check the quantities.\nSS316 template grades: nuts A4-80; bolts/washers A4-70.\nDefault washer thickness when omitted: M20 and below = 3mm; M24 and above = 4mm (washers above M20 are never thinner than 4mm).\nGaskets: write the DN and PN on each line, for example DN400 PN10 = 10. Thickness is 3mm unless you write it, for example DN400 PN10 5mm = 10. Size, standard (BS 1154 Class Y3), reinforcement and type (EPDM) follow the gasket template.\nA gasket list is read too, typed or as a photo / Excel file with the header as caption: first line such as C3D2 EPDM 3mm Thk Gasket for WDL21 Dampers (the location after "for" is the Title, or write Title:), then one row per size: 1 DN100 PN10 106.\n\nSteel work Word forms (written notification + request form): fill in this form and send it as one message:\n${steelExample}\nWrite today for the current date. Dates are day/month/year. You receive both Word files plus PDF previews rendered on the host computer. Repeating a reference number saves a new revision instead of overwriting.\n\n/id - Show your Telegram ID\n/sendemail - Email your latest order file to the administrator's mailbox (reply to one of the files from this bot to choose another order)\n/delete - If the latest file is wrong, delete that order from the records; the administrator is told by email (reply to one of the files from this bot to delete another order)\n/help - Show these instructions`); return;
  }
  const sender=[m.from.first_name,m.from.last_name].filter(Boolean).join(' ');
  const attachment=m.document || (m.photo && m.photo[m.photo.length-1]);
  if(attachment) {await handleAttachment(m,attachment,sender,uid,chat); return;}
  if(!text) {await say(chat,'Please send your order as text, or as an Excel file / photo of the table with a caption giving the project, material and Title. Send /help for the formats.');return;}
  if(isSteelForm(text)) {await handleSteel(m,text,sender,uid,chat); return;}
  // A typed gasket list (header such as C3D2 EPDM 3mm Thk Gasket for WDL21 Dampers, rows "1 DN100 PN10 106") becomes the gasket text format.
  let order, source=text;
  const gasketTable=isGasketTable(text);
  try {if(gasketTable) source=gasketOrderText(text); order=parseOrder(source,sender);}
  catch(e) {await say(chat,e.message+'\n\nPlease correct and resend the complete order. Send /help for the format.');return;}
  const job=await startJob(m,chat);
  if(job.prior) return;
  if(gasketTable) await say(chat,`Read from your message:\n${source}`.slice(0,4096));
  await deliverOrder({m,order,source,sender,uid,chat,job});
}

// One folder per Telegram message under outputs/; a message seen before is never generated twice.
async function startJob(m,chat) {
  const id=`${chat}_${m.message_id}`;
  const jobdir=path.join(root,'outputs',id);
  await fs.mkdir(jobdir,{recursive:true});
  const jobfile=path.join(jobdir,'delivery.json');
  let prior;
  try {prior=JSON.parse(await fs.readFile(jobfile,'utf8'));} catch(e) {if(e.code!=='ENOENT') throw e;}
  if(prior?.status==='sending' || prior?.status==='uncertain') await say(chat,'The previous file delivery could not be confirmed. Please check whether you received it. If not, resend the order.');
  return {id,jobdir,jobfile,prior:Boolean(prior?.status)};
}

// Generates the order file, sends it to the sender and copies a colleague's file to the administrator.
async function deliverOrder({m,order,source,sender,uid,chat,job,sourceFile}) {
  const {id,jobdir,jobfile}=job;
  await fs.writeFile(path.join(jobdir,'request.json'),JSON.stringify({telegramUserId:uid,receivedAt:new Date(m.date*1000).toISOString(),order,source,sourceFile},null,2));
  let filename;
  try {filename=await generate(order,id);}
  catch {console.error(`Excel generation failed for ${id}`);await say(chat,'Excel generation failed. Please contact the administrator. Your order has been saved on the host computer.');return;}
  await fs.writeFile(jobfile,JSON.stringify({status:'sending'}));
  let delivery;
  try {
    const sent=await sendFile(chat,filename,`${order.project} ${order.material} | Requester: ${order.requester}\n${order.sections.map(s=>s.title).join('; ').slice(0,500)}\nGenerated using the saved template standards. No extra percentage added.${sourceFile?'\nRead from your '+sourceFile+': please check the quantities against the original.':''}`);
    delivery={status:'sent',messageId:sent.message_id};
    await fs.writeFile(jobfile,JSON.stringify(delivery));
    console.log(`Order delivered ${id}`);
  } catch {
    delivery={status:'uncertain'};
    await fs.writeFile(jobfile,JSON.stringify(delivery));
    await say(chat,'The file was generated on the host computer, but Telegram delivery could not be confirmed. Please check whether you received it. If not, resend the order.');
  }
  // The administrator receives a copy of every colleague's order file as soon as it is generated.
  if(!state.owner || uid===state.owner) return;
  const who=m.from.username ? `${sender} (@${m.from.username})` : sender;
  try {const copy=await sendFile(state.owner,filename,orderNotice(order,who,delivery.status==='sent')); delivery.adminCopy='sent'; delivery.adminMessageId=copy.message_id;}
  catch(e) {console.error(`Administrator copy failed for ${id} (${e.code || 'network/local error'}). ${new Date().toISOString()}`); delivery.adminCopy='failed';}
  await fs.writeFile(jobfile,JSON.stringify(delivery));
}

// A fastener order as an Excel file or a photo of the table; the caption gives project, material and Title.
async function handleAttachment(m,attachment,sender,uid,chat) {
  const caption=(m.caption||'').trim();
  const name=attachment.file_name||'', mime=attachment.mime_type||'';
  const isPhoto=Boolean(m.photo) || /^image\//.test(mime) || /\.(png|jpe?g)$/i.test(name);
  const isXlsx=/spreadsheetml/.test(mime) || /\.xlsx$/i.test(name);
  if(!isPhoto && !isXlsx) {await say(chat,'Please send the order as text, as an Excel file (.xlsx) or as a photo of the table. Send /help for the formats.');return;}
  if((attachment.file_size||0)>20*1024*1024) {await say(chat,'The file is larger than 20 MB. Please send a smaller file.');return;}
  // No caption is needed when the table itself names the project, the grade and the title; orderTextFromTable says what is missing.
  // A gasket caption that already lists the DN / PN / Qty rows is used as typed; otherwise the table is read.
  const gasket=isGasketHeader(caption.split(/\r?\n/)[0]), fromCaption=gasket && gasketRowsIn(caption);
  const readFrom=fromCaption?'caption':isPhoto?'photo':'Excel file';
  let cfg={};
  try {cfg=JSON.parse((await fs.readFile(path.join(root,'table.json'),'utf8')).replace(/^﻿/,''));} catch(e) {if(e.code!=='ENOENT') throw e;}
  const job=await startJob(m,chat);
  if(job.prior) return;
  let text, note='';
  try {
    const source=path.join(job.jobdir,isPhoto?'source'+(name.match(/\.(png|jpe?g)$/i)?.[0]?.toLowerCase()||'.jpg'):'source.xlsx');
    await download(attachment.file_id,source);
    const xlsx=isPhoto?path.join(job.jobdir,'table.xlsx'):source;
    // Gasket photos are read by position first (numbers go to the Qty / PN heading above them, missed small figures
    // are re-read), because the cell reader loses columns of cropped screenshots; the cell reader is the fallback.
    let positionalError;
    if(gasket && isPhoto && !fromCaption) {
      try {
        const lines=await imageToLines(source,cfg);
        await fs.writeFile(path.join(job.jobdir,'ocr-lines.json'),JSON.stringify(lines));
        const grid=linesToGrid(lines);
        if(grid) text=gasketOrderText(caption,grid);
      } catch(e) {positionalError=e;}
    }
    if(!text && isPhoto && !fromCaption) await imageToXlsx(source,xlsx,cfg);
    if(gasket) {
      if(!text) {try {text=gasketOrderText(caption,fromCaption?[]:await readGrid(xlsx));} catch(e) {throw positionalError||e;}}
    }
    else {
      const table=gridToOrderLines(await readGrid(xlsx));
      text=orderTextFromTable(caption,table);
      if(table.percent) note='\n\nQuantities are used exactly as written in the table (the % allowance shown there is taken as already included). No extra 5% is added.';
    }
  } catch(e) {
    console.error(`Table reading failed for ${job.id}: ${String(e.message).split(/\r?\n/)[0]} ${new Date().toISOString()}`);
    await say(chat,`${e.message}\n\nPlease correct and resend, or send the order as text. Send /help for the format.`);
    return;
  }
  let order;
  try {order=parseOrder(text,sender);}
  catch(e) {await say(chat,`Read from your ${readFrom}:\n${text}\n\n${e.message}\n\nPlease correct and resend, or send the order as text.`);return;}
  await say(chat,`Read from your ${readFrom}:\n${text}${note}`.slice(0,4096));
  await deliverOrder({m,order,source:text,sender,uid,chat,job,sourceFile:readFrom});
}
let backoff=2000;
while(true) {
  try {
    await sendDueSummaries();
    const updates=await api('getUpdates',{offset:state.offset,timeout:45,allowed_updates:['message']});
    for (const update of updates) {
      if(update.message) await handle(update.message);
      state.offset=update.update_id+1; await save();
    }
    backoff=2000;
  } catch(e) {
    console.error(`Polling paused (${e.code || 'network/local error'}). ${new Date().toISOString()}`);
    if(e.code===401 || e.code===409) process.exit(1);
    await new Promise(resolve=>setTimeout(resolve,Math.max(backoff,(e.retryAfter||0)*1000)));
    backoff=Math.min(backoff*2,60000);
  }
}
