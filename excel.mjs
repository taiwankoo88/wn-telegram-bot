import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile, Workbook} from '@oai/artifact-tool';
const root = path.dirname(fileURLToPath(import.meta.url));
export const literal = s => /^[=+@-]/.test(s) ? "'" + s : s;

// Designation, standard, strength and remark for an item, shared by order files and weekly summaries.
export function describeItem(material, item) {
  const remark = item.parts.length > 1 ? `Combined: ${item.parts.join(' + ')}` : item.perSet || null;
  // Gasket wording follows the gasket template; the PN has its own column in the order file.
  if (item.kind === 'gasket') return {
    designation: `SIZE: DN${item.dn} ${item.thickness}MM THK Full-Face with Bolt Holes`,
    standard: 'BS 1154 Class Y3',
    grade: `PN${item.pn}`,
    remark,
  };
  const label = {bolt:'FULL THREADED HEXAGON BOLT', nut:'HEXAGON NUT',washer:'FLAT WASHER',rod:'STUD ROD'}[item.kind];
  const suffix = item.kind === 'washer' ? ` X ${item.length}mm thk.` : item.length ? `, L=${item.length}mm` : '';
  return {
    designation: `${label}: M${item.diameter}${suffix}`,
    standard: material === 'SS316' ? 'BS EN ISO 3506' : item.kind === 'washer' ? 'BS 4320' : 'BS EN ISO 4190',
    grade: material === 'HDG' ? 'G 8.8' : item.kind === 'nut' ? 'A4-80' : 'A4-70',
    remark,
  };
}

// Excel applies a shared formula to every cell in its range, even cells later filled with
// quantities or totals, so keep only the template's cached values.
function withoutSharedFormulas(wb) {
  const proto = wb.toProto();
  for (const sheet of proto.sheets) {
    sheet.sharedFormulas = [];
    for (const cell of sheet.rows.flatMap(row => row.cells ?? [])) {
      if (cell.sharedFormulaSi === undefined) continue;
      delete cell.formula; delete cell.formulaType; delete cell.formulaRef; delete cell.sharedFormulaSi;
    }
  }
  return Workbook.load(proto, {validate: true});
}

// Loads a template, names its sheet after the project and fills the project, date and requester cells.
async function open(order, template) {
  let wb = withoutSharedFormulas(await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root, 'templates', `${template}.xlsx`))));
  let s = wb.worksheets.getItemAt(0);
  if (s.name !== order.project) {
    s.name = order.project;
    // Re-import the renamed template to refresh the calculation engine's sheet registry.
    wb = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(wb));
    s = wb.worksheets.getItemAt(0);
  }
  s.getRange('B1').values = [[`TWRP ${order.project}`]];
  const date = new Intl.DateTimeFormat('en-CA', {timeZone:'Asia/Singapore',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  s.getRange('B3').values = [[new Date(date+'T00:00:00')]];
  s.getRange('B3').format.numberFormat = 'dd-mmm-yyyy';
  s.getRange('B4').values = [[literal(order.requester)]];
  return {wb, s};
}

// Recalculates, checks every quantity and formula, saves the workbook and returns its path.
// cols names the quantity, unit price and amount columns; last is the Total row.
async function save(wb, s, order, id, preview, rows, last, [qty, price, amount], name) {
  if (preview) {
    s.getRange(`${price}${rows[0].r}`).values=[[2.5]];
    if (s.getRange(`${amount}${rows[0].r}`).values[0][0] !== rows[0].qty*2.5) throw Error('Price recalculation failed');
    s.getRange(`${price}${rows[0].r}`).values=[[null]];
  }
  wb.recalculate();
  for (const row of rows) if (s.getRange(`${qty}${row.r}`).values[0][0] !== row.qty) throw Error('Quantity verification failed');
  const scan=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:10}});
  if (!scan.ndjson.includes('matched 0 entries')) throw Error('Formula check failed');
  const outdir = path.join(root,'outputs',String(id).replace(/[^a-z0-9_-]/gi,'_'));
  await fs.mkdir(outdir,{recursive:true});
  if (preview) {
    const width = amount.charCodeAt(0) - 64;
    await fs.writeFile(path.join(outdir,'preview.png'),new Uint8Array(await (await wb.render({sheetName:order.project,range:`A1:${amount}${last}`,scale:1})).arrayBuffer()));
    await fs.writeFile(path.join(outdir,'check.ndjson'),(await wb.inspect({kind:'table',sheetId:order.project,range:`A1:${amount}${last}`,include:'values,formulas',tableMaxRows:last,tableMaxCols:width,maxChars:15000})).ndjson);
  }
  const filename = path.join(outdir,name);
  await (await SpreadsheetFile.exportXlsx(wb)).save(filename);
  return filename;
}

export async function generate(order, id, preview = false) {
  if (order.material === 'EPDM') return generateGasket(order, id, preview);
  const {wb, s} = await open(order, order.material);
  s.getRange('A7:I7').copyTo(s.getRange('M1:U1'), 'all');
  s.getRange('A8:I8').copyTo(s.getRange('M2:U2'), 'all');
  s.getRange('A16:I16').copyTo(s.getRange('M3:U3'), 'all');
  s.getRange('A7:I200').unmerge();
  s.getRange('A7:I200').clear({applyTo:'all'});
  s.getRange('A4').values = [['Requester Name']];
  s.getRange('A4').format = {font:{bold:true}, wrapText:true, rowHeight:32};
  s.getRange('B4').format = {wrapText:true,horizontalAlignment:'left',font:{name:'Arial',size:12,bold:false}};
  const rows=[];
  let r=7;
  for (const [ix, sec] of order.sections.entries()) {
    s.getRange('M1:U1').copyTo(s.getRange(`A${r}:I${r}`),'all');
    s.getRange(`A${r}:I${r}`).clear({applyTo:'contents'});
    s.getRange(`A${r}:B${r}`).values = [[ix+1,literal(sec.title)]];
    s.getRange(`A${r}:G${r}`).format = {font:{name:'Arial',size:11},borders:{preset:'all',style:'thin',color:'#000000'},verticalAlignment:'center'};
    s.getRange(`B${r}`).format = {wrapText:true,font:{bold:true}};
    s.getRange(`A${r}`).format.horizontalAlignment='center';
    s.getRange(`A${r}:I${r}`).format.rowHeight = Math.max(26,Math.ceil(sec.title.length/52)*18);
    r++;
    for (const [j,item] of sec.items.entries()) {
      s.getRange('M2:U2').copyTo(s.getRange(`A${r}:I${r}`),'all');
      s.getRange(`A${r}:I${r}`).clear({applyTo:'contents'});
      const {designation, standard, grade, remark} = describeItem(order.material, item);
      s.getRange(`A${r}:I${r}`).values = [[j+1,designation,standard,order.material,grade,item.qty,remark,null,null]];
      s.getRange(`A${r}:G${r}`).format = {font:{name:'Arial',size:11},borders:{preset:'all',style:'thin',color:'#000000'},verticalAlignment:'center'};
      s.getRange(`A${r}`).format.horizontalAlignment='center';
      s.getRange(`F${r}`).format.font={bold:true};
      if (item.parts.length > 1) s.getRange(`F${r}`).formulas = [[`=SUM(${item.parts.join(',')})`]];
      s.getRange(`I${r}`).formulas = [[`=IF(H${r}="","",F${r}*H${r})`]];
      s.getRange(`F${r}`).format.numberFormat = '#,##0';
      s.getRange(`H${r}:I${r}`).format.numberFormat = '0.0000';
      s.getRange(`G${r}`).format.wrapText = true;
      s.getRange(`A${r}:I${r}`).format.rowHeight = Math.max(26,Math.ceil((remark || '').length/24)*17);
      rows.push({r,qty:item.qty}); r++;
    }
    r++;
  }
  s.getRange(`A${r}:G${r}`).merge();
  s.getRange(`A${r}`).values = [['Document Required: Mill Test Certificate (MTC) for all fastener items to be attached together with DO.']];
  s.getRange(`A${r}:G${r}`).format = {wrapText:true,rowHeight:30,borders:{preset:'outside',style:'thin',color:'#000000'}};
  r+=2;
  s.getRange(`G${r}`).values = [['Total']];
  s.getRange(`I${r}`).formulas = [[`=IF(COUNT(H8:H${r-3})=0,"",SUM(I8:I${r-3}))`]];
  s.getRange(`I${r}`).format.numberFormat='0.0000';
  s.getRange('M1:U3').clear({applyTo:'all'});
  return save(wb, s, order, id, preview, rows, r, ['F','H','I'], `${order.project} ${order.material} Fastener Order.xlsx`);
}

// Gasket orders fill the EPDM gasket template: Size, Standard, Reinforcement, PN, Type, Total Qty,
// Remarks, Unit Price and Amount in columns B-J, with each title as its own section like the fastener files.
async function generateGasket(order, id, preview) {
  let {wb, s} = await open(order, 'Gasket');
  // The template's reinforcement cell (D8) is rich text with a bold "BS EN 751". Values cannot carry
  // text runs, so every item row receives the template cell's runs through the workbook proto before
  // saving. The template's non-breaking space is replaced: the values setter rejects that string as is.
  const d8 = wb.toProto().sheets[0].rows.find(row => row.index === 8).cells.find(c => c.address === 'D8');
  const reinforcement = JSON.parse(JSON.stringify({value: d8.value, paragraphs: d8.paragraphs}).replace(/ /g, ' '));
  s.getRange('A7:J200').unmerge();
  s.getRange('A7:J200').clear({applyTo:'all'});
  const cell = {font:{name:'Arial',size:11},borders:{preset:'all',style:'thin',color:'#000000'},verticalAlignment:'center'};
  const rows=[];
  let r=7;
  for (const sec of order.sections) {
    s.getRange(`B${r}`).values = [[`Title: ${sec.title}`]];
    s.getRange(`A${r}:J${r}`).format = cell;
    s.getRange(`B${r}`).format = {wrapText:true,font:{bold:true}};
    s.getRange(`A${r}:J${r}`).format.rowHeight = Math.max(17.4,Math.ceil((7+sec.title.length)/42)*15);
    r++;
    for (const [j,item] of sec.items.entries()) {
      const {designation, standard, remark} = describeItem(order.material, item);
      s.getRange(`A${r}:H${r}`).values = [[j+1,designation,standard,reinforcement.value,item.pn,'EPDM',item.qty,remark]];
      if (item.parts.length > 1) s.getRange(`G${r}`).formulas = [[`=SUM(${item.parts.join(',')})`]];
      s.getRange(`J${r}`).formulas = [[`=IF(I${r}="","",G${r}*I${r})`]];
      s.getRange(`A${r}:J${r}`).format = {...cell,horizontalAlignment:'center',wrapText:true};
      s.getRange(`B${r}`).format.horizontalAlignment='left';
      s.getRange(`D${r}`).format.horizontalAlignment='left';
      s.getRange(`G${r}`).format = {font:{bold:true},numberFormat:'#,##0'};
      s.getRange(`I${r}:J${r}`).format.numberFormat = '#,##0.00';
      s.getRange(`A${r}:J${r}`).format.rowHeight = Math.max(32,Math.ceil((remark || '').length/20)*15);
      rows.push({r,qty:item.qty}); r++;
    }
    s.getRange(`A${r}:J${r}`).format.rowHeight = 21;
    r++;
  }
  s.getRange(`A${r}:H${r}`).merge();
  s.getRange(`A${r}`).values = [['Document Required: Certificate for all items to be attached together with DO.']];
  s.getRange(`A${r}:H${r}`).format = {font:{name:'Arial',size:11},wrapText:true,rowHeight:20,borders:{preset:'outside',style:'thin',color:'#000000'}};
  r+=2;
  s.getRange(`H${r}`).values = [['Total']];
  s.getRange(`J${r}`).formulas = [[`=IF(COUNT(I8:I${r-3})=0,"",SUM(J8:J${r-3}))`]];
  s.getRange(`H${r}:J${r}`).format = {...cell,rowHeight:21};
  s.getRange(`H${r}`).format = {font:{bold:true},horizontalAlignment:'right'};
  s.getRange(`J${r}`).format = {font:{bold:true},numberFormat:'#,##0.00'};
  // The export writes text runs only for shared-string cells (dataType 1).
  const proto = wb.toProto();
  for (const row of proto.sheets[0].rows) {
    if (!rows.some(x => x.r === row.index)) continue;
    Object.assign(row.cells.find(c => c.address === `D${row.index}`), {dataType: 1, paragraphs: structuredClone(reinforcement.paragraphs)});
  }
  wb = Workbook.load(proto, {validate: true});
  s = wb.worksheets.getItemAt(0);
  return save(wb, s, order, id, preview, rows, r, ['G','I','J'], `${order.project} EPDM Gasket Order.xlsx`);
}
