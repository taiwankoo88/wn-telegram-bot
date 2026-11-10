import fs from 'node:fs/promises';
import path from 'node:path';
import {SpreadsheetFile, Workbook} from '@oai/artifact-tool';
import {describeItem, literal} from './excel.mjs';

const DAY = 86400e3;
const SGT = 8 * 3600e3; // Singapore time has no daylight saving.
const MONTHS = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split(' ');
const WEEKDAYS = 'Sun Mon Tue Wed Thu Fri Sat'.split(' ');
const sgt = ms => new Date(ms + SGT);
const isoDate = ms => sgt(ms).toISOString().slice(0, 10);
const shortDate = ms => `${WEEKDAYS[sgt(ms).getUTCDay()]} ${sgt(ms).getUTCDate()} ${MONTHS[sgt(ms).getUTCMonth()]}`;
const weekStart = week => Date.parse(`${week}T00:00:00+08:00`);

// Weeks run Monday to Sunday in Singapore time and are named by their Monday, e.g. 2026-09-14.
export const weekOf = ms => isoDate(weekStart(isoDate(ms)) - (sgt(ms).getUTCDay() + 6) % 7 * DAY);
export const nextWeek = week => isoDate(weekStart(week) + 7 * DAY);
// A week's summary is due at 09:00 on the following Monday.
export const latestDueWeek = (now = Date.now()) => weekOf(now - 7 * DAY - 9 * 3600e3);

// Orders received during the week whose Excel file was generated (failed generations have no delivery record).
export async function loadOrders(outputs, week) {
  const from = weekStart(week);
  const entries = await fs.readdir(outputs, {withFileTypes: true}).catch(e => {
    if (e.code === 'ENOENT') return [];
    throw e;
  });
  const orders = [];
  for (const entry of entries.filter(x => x.isDirectory())) {
    const dir = path.join(outputs, entry.name);
    let request, delivery;
    try {
      request = JSON.parse(await fs.readFile(path.join(dir, 'request.json'), 'utf8'));
      delivery = JSON.parse(await fs.readFile(path.join(dir, 'delivery.json'), 'utf8'));
    } catch (e) {
      if (e.code === 'ENOENT') continue;
      throw e;
    }
    // Orders saved before receivedAt was recorded use the time request.json was written.
    const received = request.receivedAt ? Date.parse(request.receivedAt) : (await fs.stat(path.join(dir, 'request.json'))).mtimeMs;
    if (received >= from && received < from + 7 * DAY) orders.push({...request.order, received, delivered: delivery.status === 'sent'});
  }
  return orders.sort((a, b) => a.received - b.received);
}

// Title and item lines of one section, as listed in weekly summaries, administrator notices and emails.
export function sectionLines(material, section) {
  return [section.title, ...section.items.map(item => {
    const {designation, grade} = describeItem(material, item);
    // A gasket's PN is its own column in the order file, so add it here.
    return `• ${designation}${item.kind === 'gasket' ? `, ${grade}` : ''} × ${item.qty.toLocaleString('en-US')}`;
  })];
}
const orderLines = order => order.sections.flatMap(section => sectionLines(order.material, section));

// Telegram limits message and caption length; cut at a line break, since the Excel file has every item.
function fit(text, limit, more) {
  if (text.length <= limit) return text;
  const cut = text.slice(0, limit - more.length);
  return cut.slice(0, cut.lastIndexOf('\n')) + more;
}

export function summaryText(orders, week, soFar = false) {
  const from = weekStart(week), to = from + 6 * DAY;
  const year = ms => sgt(ms).getUTCFullYear();
  const range = `${shortDate(from)}${year(from) === year(to) ? '' : ` ${year(from)}`} – ${shortDate(to)} ${year(to)}`;
  const count = orders.length ? `${orders.length} order${orders.length === 1 ? '' : 's'}` : 'no orders';
  const lines = ['Weekly order summary', `${range} · ${count}${soFar ? ' so far' : ''}`];
  for (const [i, order] of orders.entries()) {
    lines.push('', `${i + 1}. ${order.project} ${order.material} · ${shortDate(order.received)} · ${order.requester}${order.delivered ? '' : ' (delivery unconfirmed)'}`, ...orderLines(order));
  }
  return fit(lines.join('\n'), 4096, '\n…the rest is in the attached Excel file.');
}

// Caption of the copy of a colleague's order file sent to the administrator (captions hold 1024 characters).
export function orderNotice(order, sender, delivered) {
  const lines = [`New order from ${sender}${delivered ? '' : ' (delivery to the sender unconfirmed)'}`, `${order.project} ${order.material} · Requester: ${order.requester}`, ...orderLines(order)];
  return fit(lines.join('\n'), 1024, '\n…see the Excel file for the full list.');
}

// Saves one row per item for the administrator's records and returns the file path.
export async function summaryWorkbook(orders, week, dir) {
  const header = ['Received', 'Requester', 'Project', 'Material', 'Title', 'Designation', 'Standard', 'Strength / PN', 'Qty', 'Remarks', 'Delivery'];
  const rows = orders.flatMap(order => order.sections.flatMap(section => section.items.map(item => {
    const {designation, standard, grade, remark} = describeItem(order.material, item);
    const received = (order.received + SGT) / DAY + 25569; // Excel date serial in Singapore time
    return [received, literal(order.requester), order.project, order.material, literal(section.title),
      designation, standard, grade, item.qty, remark, order.delivered ? 'Sent' : 'Unconfirmed'];
  })));
  const wb = Workbook.create();
  const sheet = wb.worksheets.add('Orders');
  const table = sheet.getRange(`A1:K${rows.length + 1}`);
  table.values = [header, ...rows];
  table.format = {font: {name: 'Arial', size: 10}, verticalAlignment: 'center', borders: {preset: 'all', style: 'thin', color: '#000000'}};
  sheet.getRange('A1:K1').format = {fill: '#D9D9D9', font: {bold: true}, horizontalAlignment: 'center'};
  sheet.getRange(`A2:A${rows.length + 1}`).format.numberFormat = 'dd-mmm-yyyy hh:mm';
  sheet.getRange(`I2:I${rows.length + 1}`).format.numberFormat = '#,##0';
  for (const [col, width] of [17, 16, 9, 9, 60, 44, 16, 9, 10, 22, 12].entries()) sheet.getCell(0, col).format.columnWidth = width;
  sheet.freezePanes.freezeRows(1);
  await fs.mkdir(dir, {recursive: true});
  const file = path.join(dir, `Order Summary ${week} to ${isoDate(weekStart(week) + 6 * DAY)}.xlsx`);
  await (await SpreadsheetFile.exportXlsx(wb)).save(file);
  return file;
}
