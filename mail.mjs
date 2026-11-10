import fs from 'node:fs/promises';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';
import {sectionLines} from './summary.mjs';

const root = path.dirname(fileURLToPath(import.meta.url));
const MONTHS = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split(' ');
// Singapore time, written like the Telegram messages, e.g. 15 Sep 2026 14:05.
function stamp(when) {
  const d = new Date(when.getTime() + 8 * 3600e3);
  const two = n => String(n).padStart(2, '0');
  return `${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()} ${two(d.getUTCHours())}:${two(d.getUTCMinutes())}`;
}

// Subject: project, material, requester and titles, e.g. C3D1 HDG Fasteners · Nicholas Koo · WDL12 Chemical Hanging Piperack.
export function mailSubject(order) {
  const kind = order.material === 'EPDM' ? 'EPDM Gaskets' : `${order.material} Fasteners`;
  const subject = [`${order.project} ${kind}`, order.requester, order.sections.map(s => s.title).join('; ')].join(' · ');
  // Exchange rejects subjects longer than 255 characters.
  return subject.length <= 255 ? subject : `${subject.slice(0, 254)}…`;
}

export function mailBody(order, sender, filename, when = new Date()) {
  return [
    ...order.sections.map(section => sectionLines(order.material, section).join('\n')),
    `Requester: ${order.requester} · Sent by: ${sender} (Telegram) · ${stamp(when)}\nAttachment: ${path.basename(filename)}`,
  ].join('\n\n');
}

// /delete: the order folder moves out of outputs/ into deleted/ (so it leaves /sendemail and the weekly summary,
// but is kept for the administrator) and gets a deleted.json note. Returns the new folder.
export async function archiveDeleted(dir, deletedRoot, info) {
  await fs.mkdir(deletedRoot, {recursive: true});
  let target = path.join(deletedRoot, path.basename(dir));
  if (await fs.access(target).then(() => true, () => false)) target += `-${Date.now()}`;
  await fs.rename(dir, target);
  await fs.writeFile(path.join(target, 'deleted.json'), JSON.stringify(info));
  return target;
}

// Notice to the administrator that an order file was deleted: same subject as /sendemail, marked DELETED.
export function deleteSubject(order) {
  const subject = `DELETED · ${mailSubject(order)}`;
  return subject.length <= 255 ? subject : `${subject.slice(0, 254)}…`;
}

export function deleteBody(order, sender, filename, {when = new Date(), emailedAt} = {}) {
  return [
    `${sender} (Telegram) deleted this order file on ${stamp(when)}. It is removed from the order records and the weekly summary, and cannot be emailed with /sendemail any more.`,
    ...(emailedAt ? [`This order was already emailed to you on ${stamp(new Date(emailedAt))}. Please disregard that email.`] : []),
    ...order.sections.map(section => sectionLines(order.material, section).join('\n')),
    `Requester: ${order.requester}\nDeleted file attached for your records: ${path.basename(filename)}`,
  ].join('\n\n');
}

// The order to email: the generated order whose file message was replied to (a colleague's own file, or a copy
// in the administrator's chat), otherwise the latest generated order in this chat.
export async function findOrder(outputs, {chat, owner, replyTo}) {
  const entries = await fs.readdir(outputs, {withFileTypes: true}).catch(e => {
    if (e.code === 'ENOENT') return [];
    throw e;
  });
  let latest = null;
  for (const entry of entries) {
    const [folderChat, messageId] = entry.name.split('_').map(Number);
    if (!entry.isDirectory() || !(replyTo ? folderChat === chat || chat === owner : folderChat === chat)) continue;
    const dir = path.join(outputs, entry.name);
    let request, delivery;
    try {
      request = JSON.parse(await fs.readFile(path.join(dir, 'request.json'), 'utf8'));
      delivery = JSON.parse(await fs.readFile(path.join(dir, 'delivery.json'), 'utf8'));
    } catch (e) {
      if (e.code === 'ENOENT') continue;
      throw e;
    }
    const file = (await fs.readdir(dir)).find(f => f.endsWith('.xlsx'));
    if (!file) continue;
    const found = {dir, order: request.order, file: path.join(dir, file)};
    if (replyTo) {
      if ((folderChat === chat && delivery.messageId === replyTo) || (chat === owner && delivery.adminMessageId === replyTo)) return found;
    } else if (!latest || messageId > latest.messageId) latest = {...found, messageId};
  }
  return latest;
}

// Sends through the Outlook profile on this PC. SendMail.ps1 reports whether the message left the Outbox
// (sent) or is still waiting for Outlook to connect (queued).
export async function sendMail({to, subject, body, attachment}, dir) {
  const jobFile = path.join(dir, 'email-job.json');
  await fs.writeFile(jobFile, JSON.stringify({to, subject, body, attachment}));
  const powershell = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  try {
    const {stdout} = await promisify(execFile)(powershell, ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', path.join(root, 'SendMail.ps1'), '-JobFile', jobFile], {timeout: 180000, windowsHide: true});
    return JSON.parse(stdout.trim().split(/\r?\n/).pop());
  } catch (e) {
    // execFile's message only names the command; keep the script's own error text for the log.
    const detail = [e.stderr, e.stdout].filter(Boolean).join('\n').replace(/\s+/g, ' ').trim();
    e.message = `${e.message.split(/\r?\n/)[0]}${e.killed ? ' (timed out)' : ''}${detail ? ` -- ${detail.slice(0, 600)}` : ''}`;
    throw e;
  } finally {
    await fs.rm(jobFile, {force: true});
  }
}
