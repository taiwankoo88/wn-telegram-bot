import fs from 'node:fs/promises';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';

// Steel-work Word forms (written notification + request form) for C3D1/C3D2, generated with the
// steel-forms handoff's own updater and approved templates. steel.json points at that handoff folder.
const root = path.dirname(fileURLToPath(import.meta.url));
const run = promisify(execFile);
const MONTHS = 'January February March April May June July August September October November December'.split(' ');
const SGT = 8 * 3600e3;
const TEMPLATES = {
  C3D1: {dir: 'C3D1 Forms', written: 'SS_Steel Work Written Notification C3D1_Batch_2.docx', request: 'SS_Request Form for Issuing Wirtten Notification_Batch_2 Amended.docx'},
  C3D2: {dir: 'C3D2 Forms', written: 'SS_Steel Work Written Notification C3D2.docx', request: 'SS_Request Form.docx'},
};

export const steelExample = `Project Forms: C3D1
Our Reference Number: UESH/C3D1/PM/STPR/048
Current Date: today
Description of Steel Work:
WDL12 - FA DUCTS PST TANK 1 TO 6 ( GROUND SUPPORTS ) AT +15.800 SHD
WDL15 - ROP FLUSH TANK AREA AT +14.600 SHD
Batch Number: 1
Estimated Tonnage: 5
Required Steel Delivery Date: 2/10/26`;

const FIRST_LINE = /^(C3D[12])\s*[-–,:：]?\s*(?:steel(?:\s*work)?(?:\s*forms?)?|STPR|forms?)$/i;
const LABELS = [
  ['project', /^(?:Project(?:\s*Forms?)?|Forms?)\s*[:：]\s*(.*)$/i],
  ['ref', /^(?:Our\s*)?Ref(?:erence)?(?:\s*(?:No|Number|#))?\.?\s*[:：]\s*(.*)$/i],
  ['date', /^(?:Current\s*)?Date\s*[:：]\s*(.*)$/i],
  ['desc', /^Description(?:\s*of(?:\s*the)?\s*Steel\s*Works?)?\s*[:：]\s*(.*)$/i],
  ['batch', /^Batch(?:\s*(?:No|Number|#))?\.?\s*[:：]\s*(.*)$/i],
  ['tonnage', /^(?:Estimated\s*)?Tonnage\s*[:：]\s*(.*)$/i],
  ['delivery', /^(?:Required\s*)?(?:Steel\s*)?(?:Date\s*of\s*)?Deliver(?:y|ed)?(?:\s*(?:to\s*site))?(?:\s*Date)?\s*[:：]\s*(.*)$/i],
];
const FIELD_HELP = 'Use the labels Project Forms, Our Reference Number, Current Date, Description of Steel Work, Batch Number, Estimated Tonnage and Required Steel Delivery Date. Send /help for an example.';

// A steel-work form request: the intake's "Project Forms:" line, or a first line such as "C3D1 Steel Form".
export function isSteelForm(text) {
  const first = (text.split(/\r?\n/).find(l => l.trim()) || '').trim();
  return /^\s*(?:Project(?:\s*Forms?)?|Forms?)\s*[:：]/im.test(text) || FIRST_LINE.test(first);
}

const pad2 = n => String(n).padStart(2, '0');
const ordinal = d => `${d}${10 <= d % 100 && d % 100 <= 20 ? 'th' : {1: 'st', 2: 'nd', 3: 'rd'}[d % 10] || 'th'}`;
const validDate = (y, m, d) => {
  const t = new Date(Date.UTC(y, m - 1, d));
  return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d ? t : null;
};
const monthIndex = name => MONTHS.findIndex(mn => mn.toLowerCase().startsWith(name.toLowerCase().slice(0, 3)));

// Reads a date the way update_steel_forms.py does (day first, four-digit or two-digit year, month as number or name).
function parseDate(raw) {
  const s = String(raw).trim().replace(/(\d+)(st|nd|rd|th)\b/gi, '$1');
  let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return validDate(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2}|\d{4})$/))) return validDate(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  if ((m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3,9})[\s-]+(\d{4})$/)) && monthIndex(m[2]) >= 0) return validDate(+m[3], monthIndex(m[2]) + 1, +m[1]);
  return null;
}

// The four spellings the updater writes for one date; unparseable text such as Type3-20/9/26 passes through unchanged.
export function formatDates(raw) {
  const d = parseDate(raw);
  if (!d) return {letter: String(raw).trim(), request: String(raw).trim(), long: String(raw).trim(), short: String(raw).trim()};
  const day = d.getUTCDate(), month = MONTHS[d.getUTCMonth()], year = d.getUTCFullYear();
  return {letter: `${day}-${month}-${year}`, request: `${day} ${month} ${year}`, long: `${ordinal(day)} ${month} ${year}`, short: `${day}/${d.getUTCMonth() + 1}/${String(year).slice(-2)}`};
}

export function parseSteelForm(text, now = new Date()) {
  const lines = text.replace(/&#x20;/g, ' ').replace(/\u00a0/g, ' ').split(/\r?\n/).map(l => l.replace(/\\+\s*$/, '').trim());
  if (lines.length > 60) throw Error('Maximum 60 lines per form.');
  const fields = {desc: []};
  let current = null, first = true;
  for (const line of lines) {
    if (!line || /^Please update/i.test(line)) continue;
    if (first) {
      first = false;
      const m = line.match(FIRST_LINE);
      if (m) { fields.project = m[1]; continue; }
    }
    const hit = LABELS.map(([key, re]) => [key, line.match(re)]).find(([, m]) => m);
    if (hit) {
      const [key, m] = hit, value = m[1].trim();
      current = key;
      if (key === 'desc') { if (value) fields.desc.push(value); }
      else if (value) fields[key] = value;
      continue;
    }
    if (current === 'desc') fields.desc.push(line);
    else if (current && fields[current] === undefined) fields[current] = line;
    else throw Error(`Line not recognized: ${line}\n${FIELD_HELP}`);
  }
  const project = (fields.project || '').toUpperCase().replace(/\s+/g, ' ').trim();
  if (!/^C3D[12]$/.test(project)) {
    if (/C3D1.*C3D2|C3D2.*C3D1/.test(project)) throw Error('Choose one project: write Project Forms: C3D1 or Project Forms: C3D2.');
    throw Error('Start with Project Forms: C3D1 or Project Forms: C3D2.');
  }
  const missing = [['ref', 'Our Reference Number'], ['date', 'Current Date'], ['batch', 'Batch Number'], ['tonnage', 'Estimated Tonnage'], ['delivery', 'Required Steel Delivery Date']]
    .filter(([key]) => !fields[key]).map(([, label]) => label);
  if (!fields.desc.length) missing.splice(2, 0, 'Description of Steel Work');
  if (missing.length) throw Error(`Missing: ${missing.join(', ')}. Please fill in every field.`);

  const refMatch = fields.ref.toUpperCase().replace(/\s+/g, '').match(/^(?:UESH\/)?(?:(C3D[12])\/)?(?:PM\/)?(?:STPR\/?)?(\d{1,3})$/);
  if (!refMatch) throw Error(`Reference not recognized: ${fields.ref}\nExample: UESH/${project}/PM/STPR/048`);
  if (refMatch[1] && refMatch[1] !== project) throw Error(`The reference ${fields.ref} belongs to ${refMatch[1]}, but the form is for ${project}. Please check the project or the reference.`);
  const nnn = refMatch[2].padStart(3, '0');

  let date;
  if (/^(?:today|今天)$/i.test(fields.date)) date = new Date(now.getTime() + SGT);
  else { date = parseDate(fields.date); if (!date) throw Error(`Current Date not recognized: ${fields.date}\nWrite today, or a date such as 21/9/26 or 21 September 2026.`); }
  const currentDate = `${pad2(date.getUTCDate())}-${MONTHS[date.getUTCMonth()]}-${date.getUTCFullYear()}`;

  if (!/^[A-Za-z0-9-]+$/.test(fields.batch)) throw Error(`Batch Number not recognized: ${fields.batch}. Example: Batch Number: 1`);
  const tonnage = fields.tonnage.replace(/\s+/g, ' ');
  if (!/^\d+(?:\.\d+)?(?: ?(?:t|tons?|tonnes?))?$/i.test(tonnage)) throw Error(`Estimated Tonnage not recognized: ${fields.tonnage}. Example: Estimated Tonnage: 5`);
  const delivery = fields.delivery.replace(/\s+/g, ' ');
  if (/^\d{1,2}\/\d{1,2}\/(?:\d{2}|\d{4})$/.test(delivery)) {
    if (!parseDate(delivery)) throw Error(`Required Steel Delivery Date is not a real date: ${delivery}. Use day/month/year, for example 2/10/26.`);
  } else if (!/^[A-Za-z0-9 ]{1,20}-\d{1,2}\/\d{1,2}\/(?:\d{2}|\d{4})$/.test(delivery)) {
    throw Error(`Required Steel Delivery Date not recognized: ${fields.delivery}. Use day/month/year, for example 2/10/26.`);
  }
  return {project, ref: `UESH/${project}/PM/STPR/${nnn}`, nnn, currentDate, descLines: fields.desc, batch: fields.batch, tonnage, delivery};
}

const exists = p => fs.access(p).then(() => true, () => false);

// steel.json may list several interpreters; the first one present is used.
async function pickPython(setting) {
  for (const candidate of [].concat(setting || 'python')) {
    if (!/[\\/]/.test(candidate) || await exists(candidate)) return candidate;
  }
  throw Error('No Python interpreter from steel.json was found.');
}

const powershell = () => path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const squash = s => s.replace(/\s+/g, ' ').trim();

// Runs a helper program; a failure carries the program's own last error line, not just "Command failed".
async function runTool(file, args, options) {
  try {
    return await run(file, args, {windowsHide: true, maxBuffer: 1 << 24, ...options});
  } catch (e) {
    const lines = String(e.stderr || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const detail = lines.find(l => /error|exception|not found|failed/i.test(l)) || lines.at(-1) || (e.killed ? `timed out after ${options.timeout / 1000} s` : e.message.split(/\r?\n/)[0]);
    throw Error(`${path.basename(String(args.find(a => /\.(py|ps1)$/i.test(a)) || file))}: ${detail.slice(0, 300)}`);
  }
}

// Renders both files with Word and checks what the XML cannot show: page count, signature names on
// page 1 (body text for the letter, the text box for the request form) and text-box overflow.
async function renderCheck(files, previewDir, signatures) {
  const job = path.join(previewDir, 'render-job.json');
  await fs.mkdir(previewDir, {recursive: true});
  await fs.writeFile(job, JSON.stringify({files: [files.written, files.request], pdfDir: previewDir}));
  let stdout;
  try {
    ({stdout} = await runTool(powershell(), ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', path.join(root, 'RenderDocx.ps1'), '-JobFile', job], {timeout: 180000}));
  } finally {
    await fs.rm(job, {force: true});
  }
  const records = stdout.trim().split(/\r?\n/).map(l => JSON.parse(l));
  if (!records.at(-1)?.done) throw Error('Word render did not finish');
  const byKind = {written: records.find(r => r.file === files.written), request: records.find(r => r.file === files.request)};
  const problems = [];
  for (const [kind, r] of Object.entries(byKind)) {
    if (!r) { problems.push(`${kind}: not rendered`); continue; }
    const haystack = kind === 'written' ? squash(r.page1Text) : r.textBoxes.filter(b => b.page === 1).map(b => squash(b.text)).join('\n');
    for (const name of signatures[kind] || []) if (!haystack.includes(squash(name))) problems.push(`${kind}: "${name}" is not on page 1`);
    for (const box of r.textBoxes) if (box.overflowing) problems.push(`${kind}: text box "${box.name}" overflows`);
    if (kind === 'request' && r.pages !== 1) problems.push(`request form runs to ${r.pages} pages`);
  }
  return {ok: !problems.length, problems, pages: {written: byKind.written?.pages, request: byKind.request?.pages}, pdfs: {written: byKind.written?.pdf, request: byKind.request?.pdf}};
}

// Writes the submission JSON, runs the updater into a new folder, verifies the fields and (unless told not to)
// the rendered layout. A repeated reference never overwrites an earlier submission: it gets a _rev2, _rev3 ... suffix.
export async function generateSteelForms(sub, cfg, {render = true} = {}) {
  const outputs = path.join(cfg.root, 'outputs');
  const python = await pickPython(cfg.python);
  let rev = 0, jsonFile, outDir;
  do {
    const suffix = ++rev === 1 ? '' : `_rev${rev}`;
    jsonFile = path.join(outputs, `submission_details_${sub.project}_STPR_${sub.nnn}${suffix}.json`);
    outDir = path.join(outputs, `${sub.project}_STPR_${sub.nnn}_Output${suffix}`);
  } while (await exists(jsonFile) || await exists(outDir));
  const details = {our_reference_number: sub.ref, current_date: sub.currentDate, description_of_steel_work: sub.descLines, batch_number: sub.batch, estimated_tonnage: sub.tonnage, required_steel_delivery_date: sub.delivery};
  await fs.mkdir(outputs, {recursive: true});
  await fs.writeFile(jsonFile, JSON.stringify(details, null, 2) + '\n');
  await runTool(python, [path.join(outputs, 'update_steel_forms.py'), jsonFile, '--form-set', sub.project, '--output-dir', outDir], {cwd: cfg.root, timeout: 120000});
  const docx = (await fs.readdir(outDir)).filter(f => f.endsWith('.docx'));
  const files = {written: docx.find(f => /Written Notification/i.test(f)), request: docx.find(f => /Request Form/i.test(f))};
  if (docx.length !== 2 || !files.written || !files.request) throw Error(`Expected the two Word files in ${outDir}, found: ${docx.join(', ') || 'none'}`);
  for (const k in files) files[k] = path.join(outDir, files[k]);

  const t = TEMPLATES[sub.project];
  const dates = formatDates(sub.currentDate), delivery = formatDates(sub.delivery);
  const job = path.join(outDir, 'verify-job.json');
  await fs.writeFile(job, JSON.stringify({
    templates: {written: path.join(outputs, t.dir, t.written), request: path.join(outputs, t.dir, t.request)},
    outputs: files,
    expected: {ref: sub.ref, letterDate: dates.letter, requestDate: dates.request, longDelivery: delivery.long, shortDelivery: delivery.short, descLines: sub.descLines, batch: sub.batch, tonnage: sub.tonnage},
  }));
  let verify;
  try {
    verify = JSON.parse((await runTool(python, [path.join(root, 'steel_verify.py'), job], {timeout: 60000})).stdout);
  } finally {
    await fs.rm(job, {force: true});
  }
  let layout = null;
  if (render && verify.ok) {
    try { layout = await renderCheck(files, path.join(outDir, 'preview'), verify.signatures); }
    catch (e) { layout = {ok: false, unavailable: true, problems: [String(e.message).split(/\r?\n/)[0]]}; }
  }
  return {rev, jsonFile, outDir, files, verify, layout};
}

// Caption for the delivered files: what was filled in and what was checked (Telegram captions hold 1024 characters).
export function steelNotice(sub, result) {
  const dates = formatDates(sub.currentDate), delivery = formatDates(sub.delivery);
  const tonnage = /[a-z]/i.test(sub.tonnage) ? sub.tonnage : `${sub.tonnage} t`;
  const lines = [
    `${sub.project} · ${sub.ref} · Batch ${sub.batch} · ${tonnage}`,
    `Date ${dates.letter} · Delivery ${delivery.long}`,
    ...sub.descLines,
  ];
  const checks = ['fields verified in both Word files'];
  const layout = result.layout;
  if (!layout) checks.push('layout not visually checked, please check page 1 when opening');
  else if (layout.unavailable) checks.push('layout not visually checked (Word unavailable), please check page 1 when opening');
  else if (layout.ok) checks.push(`Word layout OK: ${layout.pages.written} page letter, ${layout.pages.request} page request form, all signature names on page 1`);
  else checks.push(`LAYOUT CHECK FAILED: ${layout.problems.join('; ')}`);
  lines.push(`Checks: ${checks.join(' · ')}`);
  lines.push(`Saved: ${path.basename(result.outDir)}${result.rev > 1 ? ` (revision ${result.rev}; the earlier submission with this reference was kept)` : ''}`);
  const text = lines.join('\n');
  if (text.length <= 1024) return text;
  const more = '\n…see the Word files for the full description.';
  const cut = text.slice(0, 1024 - more.length);
  return cut.slice(0, cut.lastIndexOf('\n')) + more;
}
