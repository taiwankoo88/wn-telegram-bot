import fs from 'node:fs/promises';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

// Fastener orders sent as an Excel file or a photo of a table. Two layouts are read:
// - columns: sizes down the first column (M24X210, M20x60 (Bolt only), M16X2m ...) and Bolt / Nut /
//   Washer / Stud rod quantity columns across the top;
// - list: one item per row described in words (HEXAGON BOLT: M12 Gr 8.8, L=45mm / NUT: M12 /
//   WASHER: M12 X 3mm THKS) with a quantity column, often under a title row such as
//   "C3D2 | C3D2 WDL23 Cable Rack Type 1-4 | 5%".
// Either way the table becomes the text order format, so parsing, consolidation and the Excel
// template work exactly as for a typed order. Photos go through the PDF bot's offline OCR converter.
const run = promisify(execFile);
const exists = p => fs.access(p).then(() => true, () => false);
const clean = c => String(c ?? '').replace(/ /g, ' ').trim();

// First sheet of a workbook as a grid of trimmed strings (numbers become their digits).
export async function readGrid(xlsxPath) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(xlsxPath));
  const values = wb.worksheets.getItemAt(0).getUsedRange(true).values || [];
  return values.map(row => row.map(clean));
}

const KINDS = [['rod', /stud|\brod\b/i], ['washer', /washer/i], ['nut', /\bnuts?\b/i], ['bolt', /\bbolts?\b/i]];
// A worded item description: the item type comes before the size (HEXAGONBOLT:M10, NUT: M12, WASHER:M20X3mm).
// The item word may come before or after the size (HEXAGON BOLT M10, M16 Nut, M16Nut).
const DESCRIPTION = /(?:bolt|nut|washer|stud|rod)[^]*?M\s?\d{1,2}(?![\d.])|M\s?\d{1,2}(?![\d.])[^]*?(?:bolt|nut|washer|stud|rod)/i;
const PERCENT = /\b\d+(?:\.\d+)?\s*%/;

// SS316 or HDG as named or implied (A4-70 / A4-80 grades, Gr 8.8) by a piece of table text.
function materialOf(text) {
  if (/\b(?:SS\s?316L?|SS16)\b|stainless|A4\s*[-–]?\s*[78]0/i.test(text)) return 'SS316';
  if (/\bHDG\b|galvani[sz]ed|Gr(?:ade)?\.?\s*8\.8|\b8\.8\b/i.test(text)) return 'HDG';
  return null;
}

// OCR drops the spaces of long merged cells (C3D2WDL23CableRackType1-4); put them back.
function respace(title) {
  if (/\s/.test(title) || title.length < 12) return title;
  return title.replace(/(?<=[a-z])(?=[A-Z\d])|(?<=\d)(?=[A-Z][a-z])/g, ' ');
}

// Project and title from the rows above the table header (or above the first item).
function titleInfo(rows) {
  const text = rows.flat().filter(Boolean).join(' ');
  const project = text.match(/C3D[12]/i)?.[0].toUpperCase() || null;
  const percent = PERCENT.test(text);
  let title = rows.flat().map(c => c.replace(PERCENT, '').trim()).filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
  for (let before = ''; before !== title;) { before = title; title = title.replace(/^C3D[12]\b\s*[-–:：|/]?\s*|^C3D[12](?=[A-Z])/i, '').trim(); }
  title = title ? respace(title) : null;
  // An all-capitals title that OCR ran together (MATERIALLISTFORBIOSORPTIONBLOWERS) cannot be split; the caption must give it.
  if (title && !/\s/.test(title) && title.length > 25) title = null;
  return {project, title, percent};
}

// Text order lines for the table, plus the project, material and title named in it (if any) so the
// caption can be checked or completed.
export function gridToOrderLines(grid) {
  const rows = joinSizeColumns(grid.map(r => r.map(clean)).filter(r => r.some(Boolean)));
  // A column heading names one item type; a cell naming several (Bolt, Nut & Washer A4-70) is a section label of a list.
  const header = rows.findIndex(r => r.some(c => c.length <= 30 && KINDS.filter(([, re]) => re.test(c)).length === 1) && !r.some(c => /^M\s*\d+/i.test(c) || DESCRIPTION.test(c)));
  const columns = header >= 0 && rows.slice(header + 1).some(r => r.some(c => /^M\s*\d+/i.test(c)));
  const result = columns ? fromColumns(rows, header) : fromList(rows);
  if (result.problems.length) throw Error(result.problems.slice(0, 6).join('\n'));
  if (!result.lines.length) throw Error('No bolts, nuts, washers or stud rods with quantities were found in the table.');
  delete result.problems;
  return result;
}

// Item lists with separate Size and Length columns (DESCRIPTION | SIZE | LENGTH (MM) | QTY, added 2026-10-09):
// the size and length are added to the description (BS 4190 BOLT A4-70 M27 L=145mm) so the list reader can read it.
function joinSizeColumns(rows) {
  const h = rows.findIndex(r => r.some(c => /^size\b/i.test(c)) && r.some(c => /qty|quantity|数量/i.test(c)));
  if (h < 0) return rows;
  const sizeCol = rows[h].findIndex(c => /^size\b/i.test(c));
  const lengthCol = rows[h].findIndex(c => /length|长度/i.test(c));
  return rows.map((r, i) => {
    if (i <= h || !/^M\s?\d{1,2}$/i.test(r[sizeCol] || '')) return r;
    const desc = r.findIndex((c, j) => j !== sizeCol && /bolt|nut|washer|stud|rod/i.test(c) && !/M\s?\d/i.test(c));
    if (desc < 0) return r;
    const out = [...r];
    const length = lengthCol >= 0 ? out[lengthCol].match(/^(\d+(?:\.\d+)?)\s*(?:mm)?$/i)?.[1] : null;
    out[desc] = `${out[desc]} ${out[sizeCol].replace(/\s/g, '')}${length ? ` L=${length}mm` : ''}`;
    out[sizeCol] = '';
    if (lengthCol >= 0) out[lengthCol] = '';
    return out;
  });
}

function fromColumns(rows, headerIndex) {
  const columns = [];
  rows[headerIndex].forEach((cell, index) => {
    const kind = KINDS.find(([, re]) => re.test(cell))?.[0];
    if (kind) columns.push({index, kind, thickness: kind === 'washer' ? cell.match(/(\d+(?:\.\d+)?)\s*mm/i)?.[1] : null});
  });
  const material = materialOf(rows.slice(0, headerIndex + 1).flat().join(' '));
  const lines = [], problems = [];
  for (const row of rows.slice(headerIndex + 1)) {
    const size = row.find(c => /^M\s*\d+/i.test(c));
    if (!size) {
      if (!/total|sum|remark/i.test(row.join(' ')) && row.some(c => /\d/.test(c))) problems.push(`Row not recognized: ${row.filter(Boolean).join(' | ')}`);
      continue;
    }
    const m = size.match(/^M\s*(\d+)\s*(?:[xX×*]\s*(\d+(?:\.\d+)?)\s*(mm|m)?\b)?/);
    const diameter = m[1], length = m[2], unit = m[3]?.toLowerCase() || 'mm';
    for (const col of columns) {
      const raw = (row[col.index] || '').replace(/,/g, '').replace(/\s+/g, '');
      if (!raw || raw === '-') continue;
      if (!/^\d+$/.test(raw)) { problems.push(`Quantity not recognized for ${size}: ${row[col.index]}`); continue; }
      if (Number(raw) === 0) continue;
      if (col.kind === 'bolt' || col.kind === 'rod') {
        if (!length) { problems.push(`${size}: a ${col.kind === 'rod' ? 'stud rod' : 'bolt'} needs a length, for example M24x210`); continue; }
        lines.push(`M${diameter}x${length}${unit} ${col.kind === 'rod' ? 'stud rod' : 'bolt'} = ${raw}`);
      } else if (col.kind === 'nut') lines.push(`M${diameter} nut = ${raw}`);
      else lines.push(`M${diameter}${col.thickness ? `x${col.thickness}mm` : ''} washer = ${raw}`);
    }
  }
  return {lines, problems, material, ...titleInfo(rows.slice(0, headerIndex))};
}

function fromList(rows) {
  const descriptionCell = r => r.findIndex(c => DESCRIPTION.test(c));
  const first = rows.findIndex(r => descriptionCell(r) >= 0);
  if (first < 0) throw Error('No table header with Bolt, Nut or Washer columns, and no item descriptions such as HEXAGON BOLT: M12 Gr 8.8, L=45mm, were found.');
  // The quantity column: a heading with Qty / Quantity (the Total one if there are several).
  let headerIndex = -1, qtyColumn = -1;
  for (let i = first - 1; i >= 0 && headerIndex < 0; i--) {
    const qty = rows[i].map((c, index) => ({c, index})).filter(({c}) => /qty|quantity|数量|\bpcs\b/i.test(c));
    if (qty.length) { headerIndex = i; qtyColumn = (qty.find(({c}) => /total/i.test(c)) || qty.at(-1)).index; }
  }
  const lines = [], problems = [], materials = new Set();
  for (const row of rows.slice(first)) {
    const at = descriptionCell(row);
    if (at < 0) {
      // Section labels such as "Bolt, Nut & Washer A4-70" have no size; only a row with a size is an unread item.
      if (!/total|sum|remark/i.test(row.join(' ')) && row.some(c => /M\s?\d{1,2}(?![\d.])/i.test(c))) problems.push(`Row not recognized: ${row.filter(Boolean).join(' | ')}`);
      continue;
    }
    // Ignore a translation in brackets, e.g. （六角螺栓：M12不锈钢8.8级，长=45mm）.
    let text = row[at].replace(/\s*[（(][^（(]*[㐀-鿿][^]*$/, '').trim();
    // OCR sometimes merges a whole row into one cell: "25b Flat Washer M24 (... 4 thk) 20 Nos".
    const inline = text.match(/\s(\d[\d,]*)\s*(?:nos|no|ns|pcs)\.?\s*$/i);
    if (inline) text = text.slice(0, inline.index).replace(/^\d+[a-z]?\s+(?=\D)/i, '').trim();
    const shown = text.replace(/\s+/g, ' ');
    // The Qty heading counts only when it lies right of the description; OCR can merge the whole heading
    // row into one cell ("Item Description Qty Unit"), and then the last number after the description is used.
    const raw = (qtyColumn > at ? row[qtyColumn] || '' : row.slice(at + 1).filter(c => /^[\d,\s]+$/.test(c)).at(-1) || inline?.[1] || '').replace(/[,\s]/g, '');
    if (!raw || raw === '-') { problems.push(`No quantity for ${shown}`); continue; }
    if (!/^\d+$/.test(raw)) { problems.push(`Quantity not recognized for ${shown}: ${row[qtyColumn]}`); continue; }
    if (Number(raw) === 0) continue;
    // OCR runs words together (BS4190NUTSA4-70), so NUT needs no word break after it.
    const found = [/stud|rod\b/i, /washer/i, /nut/i, /bolt/i].map(re => re.test(text));
    const kind = found[0] ? 'rod' : found.slice(1).filter(Boolean).length > 1 ? null : ['washer', 'nut', 'bolt'][found.slice(1).indexOf(true)];
    if (!kind) { problems.push(`One row names more than one item, please list them separately: ${shown}`); continue; }
    const diameter = text.match(/M\s?(\d{1,2})(?![\d.])/i)[1];
    const grade = text.match(/Gr(?:ade)?\.?\s*(\d+\.\d)/i)?.[1];
    if (grade && grade !== '8.8') { problems.push(`Grade ${grade} is not in the templates (HDG is G 8.8): ${shown}`); continue; }
    // The grade column (SS A4-70, SS316L, Gr 8.8) names the material, so a caption that disagrees is refused.
    const material = materialOf(row.join(' '));
    if (material) materials.add(material);
    // OCR often merges words (M10X3mmTHKS, L=50mm）), so units are not followed by a word boundary.
    const times = text.match(/M\s?\d{1,2}\s*[xX×*]\s*(\d+(?:\.\d+)?)\s*(mm|m)?/i);
    if (kind === 'bolt' || kind === 'rod') {
      const length = text.match(/\bL(?:ength)?\s*[=:：]\s*(\d+(?:\.\d+)?)\s*(mm|m)?/i) || times;
      if (!length) { problems.push(`${shown}: a ${kind === 'rod' ? 'stud rod' : 'bolt'} needs a length, for example L=45mm`); continue; }
      lines.push(`M${diameter}x${length[1]}${length[2]?.toLowerCase() || 'mm'} ${kind === 'rod' ? 'stud rod' : 'bolt'} = ${raw}`);
    } else if (kind === 'nut') lines.push(`M${diameter} nut = ${raw}`);
    else {
      const thickness = (times && /mm/i.test(times[2] || '') ? times : null) || text.match(/(\d+(?:\.\d+)?)\s*(?:mm)?\s*(?:THK|THICK|T\b)/i) || text.match(/THK\.?\s*(\d+(?:\.\d+)?)/i);
      lines.push(`M${diameter}${thickness ? `x${thickness[1]}mm` : ''} washer = ${raw}`);
    }
  }
  if (materials.size > 1) problems.push('The table mixes SS316 and HDG items. Please send each material separately.');
  // Without a grade on the item rows, a section label above them (Bolt, Nut & Washer A4-70) names the material.
  return {lines, problems, material: [...materials][0] || materialOf(rows.slice(0, first).flat().join(' ')), ...titleInfo(rows.slice(0, headerIndex >= 0 ? headerIndex : first))};
}

// The caption gives the project, material and Title; whatever it leaves out is taken from the table
// (project and title row, grade of the items). The table supplies the items.
export function orderTextFromTable(caption, table) {
  const head = String(caption || '').replace(/\s*[/|]?\s*(?=(?:Title|Location|Request[eo]r(?:\s*Name)?)\s*[:：])/gi, '\n')
    .split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let project = null, material = null;
  const rest = [];
  for (const line of head) {
    const m = line.match(/^(?:(C3D[12])\s*[,，:：/-]?\s*)?(?:(SS316|SS16|HDG)\s*)?(?:fasteners?)?\s*[:：]?$/i);
    if (m && (m[1] || m[2])) { project ||= m[1]?.toUpperCase() || null; material ||= m[2]?.toUpperCase().replace(/^SS16$/, 'SS316') || null; }
    else rest.push(line);
  }
  if (project && table.project && project !== table.project) throw Error(`The table says ${table.project} but the caption says ${project}. Please check and resend.`);
  if (material && table.material && table.material !== material) throw Error(`The table says ${table.material} but the caption says ${material}. Please check and resend.`);
  project ||= table.project;
  material ||= table.material;
  if (!project || !material) throw Error(`Write the project and material on the first line of the caption, for example ${project || 'C3D1'} ${material || 'SS16'}, then Title: and the location.`);
  const hasTitle = rest.some(l => /^(?:Title|Location|标题|位置)\s*[:：]\s*\S/i.test(l) || /^WDL\d+\s+\S/i.test(l));
  if (!hasTitle && !table.title) throw Error('Add Title: followed by the location to the caption, for example Title: Knife Gate Valves.');
  return [`${project} ${material}`, ...rest, ...(hasTitle ? [] : [`Title: ${table.title}`]), ...table.lines].join('\n');
}

// Gasket orders as a table (added 2026-10-09): a header line such as "C3D2 Fibre Reinforced EPDM 3mm Thk Gasket
// for WDL21 Dampers", then one DN / PN / Qty row per size ("1 DN100 PN10 106"), typed in the message or caption,
// or in a photo or Excel file. The rows become "DN100 PN10 3mm = 106" lines for the gasket parser.
const STRICT_GASKET_HEAD = /^C3D[12]\s*[,，:：-]?\s*(?:EPDM(?:\s*GASKETS?)?|GASKETS?(?:\s*EPDM)?)\s*[:：]?$/i;
export const isGasketHeader = line => /gasket|epdm/i.test(line || '');
// A typed gasket list that the text parser does not read as it stands (header other than "C3D1 Gasket").
export const isGasketTable = text => {
  const first = String(text || '').trim().split(/\r?\n/)[0].trim();
  return isGasketHeader(first) && !STRICT_GASKET_HEAD.test(first);
};
export const gasketRowsIn = text => String(text || '').split(/\r?\n/).slice(1).some(l => /\bDN\s?\d+/i.test(l));

export function gasketOrderText(caption, grid = []) {
  const lines = String(caption || '').split(/\r?\n/).map(l => clean(l)).filter(Boolean);
  const head = lines[0] || '';
  const project = head.match(/C3D[12]/i)?.[0].toUpperCase();
  if (!project) throw Error('Write the project on the first line, for example C3D2 EPDM 3mm Thk Gasket for WDL21 Dampers.');
  const other = head.match(/\b(?:neoprene|nitrile|NBR|viton|FKM|PTFE|CNAF|graphite|silicone|spiral\s*wound|cork)\b/i);
  if (other) throw Error(`Only EPDM gaskets are accepted: ${head}`);
  const thickness = head.match(/(\d+(?:\.\d+)?)\s*mm/i)?.[1];
  const rest = lines.slice(1).filter(l => !/\bDN\s?\d+/i.test(l));
  const requester = rest.filter(l => /^(?:Request[eo]r(?:\s*Name)?|申请人)\s*[:：]/i.test(l));
  const titleLine = rest.find(l => /^(?:Title|Location|标题|位置)\s*[:：]\s*\S/i.test(l));
  const title = titleLine ? titleLine.replace(/^[^:：]+[:：]\s*/, '') : head.match(/\bfor\s+(.+)$/i)?.[1].trim();
  if (!title) throw Error('Add Title: followed by the location, for example Title: WDL21 Dampers.');
  const rows = gasketRowsIn(caption) ? lines.slice(1).filter(l => /\bDN\s?\d+/i.test(l)).map(l => [l]) : grid;
  // Tables with their own Qty and PN columns (Item | Description | Qty | PN, PN written as a bare 10; added 2026-11-06):
  // the values are taken from those columns, so a number is never mistaken for the other column.
  const headerRow = rows.findIndex(r => r.some(c => /^PN$/i.test(clean(c))) || r.some(c => /^(?:qty|quantity)\b/i.test(clean(c))) && !r.some(c => /DN\s?\d/i.test(c)));
  const pnCol = headerRow >= 0 ? rows[headerRow].findIndex(c => /^PN$/i.test(clean(c))) : -1;
  const qtyCol = headerRow >= 0 ? rows[headerRow].findIndex(c => /^(?:qty|quantity)\b/i.test(clean(c))) : -1;
  // "Gasket 3mm EPDM" written above the rows gives the thickness when the caption does not.
  const tableThk = thickness ? null : rows.map(r => r.map(clean).join(' ')).find(t => /gasket|epdm/i.test(t) && !/DN\s?\d/i.test(t) && /\d\s*mm/i.test(t))?.match(/(\d+(?:\.\d+)?)\s*mm/i)?.[1];
  const items = [], problems = [];
  for (const [index, row] of rows.entries()) {
    if (index <= headerRow) continue;
    // OCR merges the row number into the DN cell (2DN150, 1|DN100, LDN800), so DN and PN need no word break before them.
    const columnValue = col => col >= 0 && col < row.length ? clean(row[col]).replace(/\s+/g, '') : '';
    const text = row.map(clean).filter((c, i) => c && i !== pnCol && i !== qtyCol).join(' ').replace(/(DN|PN)\s+(?=\d)/gi, '$1');
    const dn = text.match(/DN(\d+)\b/i);
    if (!dn) {
      if (/PN\d|\bclass\s*\d/i.test(text)) problems.push(`Row not recognized: ${text}`);
      continue;
    }
    const pnCell = columnValue(pnCol).match(/^(?:PN)?(\d+(?:\.\d+)?)$/i);
    const pn = text.match(/PN(\d+(?:\.\d+)?)\b/i) || pnCell;
    const shown = [text, columnValue(qtyCol), columnValue(pnCol) && `PN ${columnValue(pnCol)}`].filter(Boolean).join(' ');
    if (!pn) { problems.push(`DN${dn[1]} has no PN rating. Only EPDM gaskets with a PN rating are accepted: ${shown}`); continue; }
    const after = pn === pnCell ? text.slice(dn.index + dn[0].length) : text.slice(Math.max(dn.index + dn[0].length, pn.index + pn[0].length));
    const qty = qtyCol >= 0 ? columnValue(qtyCol).match(/^\d{1,3}(?:,\d{3})+$|^\d+$/)?.[0] : after.split(/\s+/).filter(t => /^\d{1,3}(?:,\d{3})+$|^\d+$/.test(t)).at(-1);
    if (!qty) { problems.push(`No quantity for DN${dn[1]} PN${pn[1]}`); continue; }
    const rowThk = after.match(/(\d+(?:\.\d+)?)\s*mm/i)?.[1] || thickness || tableThk;
    if (Number(qty.replace(/,/g, '')) === 0) continue;
    items.push(`DN${dn[1]} PN${pn[1]}${rowThk ? ` ${rowThk}mm` : ''} = ${qty.replace(/,/g, '')}`);
  }
  if (problems.length) throw Error(problems.slice(0, 6).join('\n'));
  if (!items.length) throw Error('No gasket sizes with DN, PN and quantity were found. Example row: 1 DN100 PN10 106');
  return [`${project} Gasket`, ...requester, `Title: ${title}`, ...items].join('\n');
}

// A photo read as positioned text lines (ocr_lines.py, beside this file), for gasket tables whose cell borders the
// workbook converter misses. Small figures missing under a Qty / PN heading are re-read from a close crop.
export async function imageToLines(imagePath, cfg) {
  const script = cfg.image2xlsx;
  if (!script || !(await exists(script))) throw Error('table.json does not name a usable image2xlsx.py');
  let python = 'python';
  for (const candidate of [].concat(cfg.python || 'python')) if (!/[\\/]/.test(candidate) || await exists(candidate)) { python = candidate; break; }
  const helper = fileURLToPath(new URL('./ocr_lines.py', import.meta.url));
  const {stdout} = await run(python, [helper, script, imagePath], {cwd: path.dirname(script), windowsHide: true, timeout: 240000, maxBuffer: 1 << 22, env: {...process.env, PYTHONIOENCODING: 'utf-8'}});
  return JSON.parse(stdout.trim().split(/\r?\n/).pop());
}

// Positioned lines as a grid: the line with a Qty / PN heading gives the columns, and every later word goes to the
// heading nearest to it. Lines above the heading stay single-cell title rows. Null when there is no such heading.
export function linesToGrid({lines}) {
  const heading = /^(?:qty|quantity|pn)$/i;
  const h = lines.findIndex(line => line.some(i => heading.test(i.text.trim())));
  if (h < 0) return null;
  const centres = lines[h].map(i => (i.x0 + i.x1) / 2);
  const grid = lines.slice(0, h).map(line => [line.map(i => i.text).join(' ')]);
  grid.push(lines[h].map(i => i.text));
  for (const line of lines.slice(h + 1)) {
    const row = centres.map(() => '');
    for (const i of line) {
      const c = (i.x0 + i.x1) / 2;
      const col = centres.reduce((best, x, n) => Math.abs(x - c) < Math.abs(centres[best] - c) ? n : best, 0);
      row[col] = row[col] ? `${row[col]} ${i.text}` : i.text;
    }
    grid.push(row);
  }
  return grid;
}

// Reads a photo of a table into a workbook with the PDF bot's offline OCR converter (table.json names it).
export async function imageToXlsx(imagePath, xlsxPath, cfg) {
  const script = cfg.image2xlsx;
  if (!script || !(await exists(script))) throw Error('table.json does not name a usable image2xlsx.py');
  let python = 'python';
  for (const candidate of [].concat(cfg.python || 'python')) if (!/[\\/]/.test(candidate) || await exists(candidate)) { python = candidate; break; }
  try {
    await run(python, [script, imagePath, xlsxPath], {cwd: path.dirname(script), windowsHide: true, timeout: 240000, maxBuffer: 1 << 22});
  } catch (e) {
    const out = `${e.stdout || ''}\n${e.stderr || ''}`;
    if (/NO_TABLES/.test(out)) throw Error('No table with ruled lines was found in the photo. Please send a clearer photo or the Excel file.');
    throw Error(`Photo reading failed: ${(out.trim().split(/\r?\n/).filter(Boolean).pop() || e.message).slice(0, 200)}`);
  }
}
