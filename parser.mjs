export const example = `C3D1 SS316
Requester Name: Nicholas
Title: WDL13 Membrane Tank Pipe Header Valve Connection
M16x110mm bolt = 2880
M16 nut = 2880
M16x130mm bolt = 5760
M16 nut = 5760
M16x3mm flat washer = 17280`;

export const gasketExample = `C3D1 Gasket
Requester Name: Nicholas
Title: WDL12 Pump Gallery 2 and 3 PSD
DN400 PN10 = 10
DN200 PN10 = 30`;

export function parseOrder(text, requester = '') {
  const lines = text.replace(/&#x20;/g, ' ').replace(/ /g, ' ').replace(/[×＊]/g, 'x').split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  if (lines.length > 100) throw Error('Maximum 100 lines per order. Please split this into smaller orders.');
  // Gasket orders start with C3D1 Gasket (EPDM, EPDM Gasket and Gasket EPDM are also accepted) and use the EPDM gasket template.
  const head = lines.shift()?.match(/^(C3D[12])\s*[,，:：-]?\s*(SS316|SS16|HDG|EPDM(?:\s*GASKETS?)?|GASKETS?(?:\s*EPDM)?)\s*[:：]?$/i);
  if (!head) throw Error('Start with C3D1 SS316, C3D2 SS316, C3D1 HDG, C3D2 HDG, C3D1 Gasket or C3D2 Gasket on the first line.');
  const material = /^SS/i.test(head[2]) ? 'SS316' : /^HDG$/i.test(head[2]) ? 'HDG' : 'EPDM';
  const order = {project: head[1].toUpperCase(), material, requester, sections: []};
  let section;
  for (const original of lines) {
    const req = original.match(/^(?:Request[eo]r(?:\s*Name)?|申请人)\s*[:：]\s*(.+)$/i);
    if (req) { order.requester = req[1].trim(); continue; }
    const title = original.match(/^(?:Title|Location|标题|位置)\s*[:：]\s*(.+)$/i) || original.match(/^(WDL\d+\s+.+)$/i);
    if (title) {
      if (title[1].length > 160) throw Error('Please keep each location title within 160 characters.');
      section = {title: title[1], sets: 1, items: []}; order.sections.push(section); continue;
    }
    if (!section) throw Error('Add Title: followed by the location name before listing items.');
    const sets = original.match(/^Sets\s*[:：=]\s*(\d+)$/i);
    if (sets) {
      if (section.items.length || section.sets !== 1) throw Error('Place Sets: followed by the number immediately after the Title, before any items.');
      section.sets = positive(sets[1]); continue;
    }
    // Read quantity only after an explicit delimiter (= 70, : 70, qty 70, qty=70, qty: 70); never infer it from size or length.
    const qtyMatch = original.match(/\s*(?:(?:\b(?:qty|quantity)\b\.?|数量)\s*[:：=]?|=|[:：]|\s-\s)\s*([\d,]+)\s*(?:pcs|nos|pieces|个)?\s*$/i);
    if (!qtyMatch) throw Error(`Quantity not recognized: ${original}\nExample: ${order.material === 'EPDM' ? 'DN400 PN10 = 10' : 'M16x65mm bolt = 180'}.`);
    const body = original.slice(0, qtyMatch.index).trim();
    const qty = positive(qtyMatch[1]);
    const item = order.material === 'EPDM' ? gasketItem(body, original) : fastenerItem(body, original, order.material);
    const total = qty * section.sets;
    positive(String(total));
    const previous = section.items.find(x => ['kind', 'diameter', 'length', 'dn', 'pn', 'thickness'].every(k => x[k] === item[k]));
    if (previous) { previous.parts.push(total); previous.qty = positive(String(previous.qty + total)); }
    else section.items.push({...item, qty: total, parts: [total], perSet: section.sets > 1 ? `${section.sets} sets x ${qty}` : ''});
  }
  if (!order.requester.trim()) throw Error('Please enter Requester Name: followed by your name.');
  if (order.requester.length > 100) throw Error('Please keep Requester Name within 100 characters.');
  if (!order.sections.length || order.sections.some(s => !s.items.length)) throw Error('Each Title must have at least one item.');
  return order;
}
// Bolt, nut, flat washer or stud rod. length is the bolt or rod length, or the washer thickness, in mm.
function fastenerItem(body, original, material) {
  if (/\b(?:gaskets?|EPDM)\b/i.test(body)) throw Error(`Gaskets are not accepted in a fastener order: ${original}. Send them as a separate order starting with C3D1 Gasket or C3D2 Gasket.`);
  const mats = body.match(/\b(?:SS316|SS16|HDG|SS)\b/gi) || [];
  for (const mat of mats) if ((/^SS/i.test(mat) ? 'SS316' : 'HDG') !== material) throw Error(`Material conflict: ${original}. Send different materials in separate orders.`);
  body = body.replace(/\b(?:SS316|SS16|HDG|SS)\b/gi, '').trim();
  const kindMatch = body.match(/\b(stud\s*rod|threaded\s*rod|flat\s*washer|washer|hex(?:agon)?\s*bolt|bolt|hex(?:agon)?\s*nut|nut)\b/i);
  if (!kindMatch) throw Error(`Item type not recognized: ${original}. Specify bolt, nut, washer or stud rod.`);
  const kind = /rod/i.test(kindMatch[0]) ? 'rod' : /washer/i.test(kindMatch[0]) ? 'washer' : /nut/i.test(kindMatch[0]) ? 'nut' : 'bolt';
  body = body.replace(kindMatch[0], '').trim();
  body = body.replace(/\b(?:length|L)\s*=?\s*/i, 'x');
  const size = body.match(/^M\s*(\d+)\s*(?:x\s*(\d+(?:\.\d+)?)\s*(mm|m)?\s*(?:thk\.?)?)?\s*$/i);
  if (!size) throw Error(`Size format not recognized: ${original}. Example: M16x65mm bolt = 180.`);
  const diameter = positive(size[1]);
  let length = size[2] ? Number(size[2]) * (size[3]?.toLowerCase() === 'm' ? 1000 : 1) : null;
  if (kind === 'nut' && length !== null) throw Error(`A nut does not require a length: ${original}`);
  if ((kind === 'bolt' || kind === 'rod') && (!length || !Number.isInteger(length))) throw Error(`Specify a length in whole millimetres: ${original}`);
  // Washer thickness: M20 and below 3 mm; M24 and above at least 4 mm (user rule, 2026-10-09).
  if (kind === 'washer' && length === null) length = diameter <= 20 ? 3 : 4;
  if (kind === 'washer' && (!length || length > 20)) throw Error('Specify the washer thickness. Example: M16x3mm washer = 200.');
  if (kind === 'washer' && diameter > 20 && length < 4) throw Error(`Washers above M20 must be at least 4mm thick: ${original}`);
  if (kind === 'rod' && material === 'HDG') throw Error('The HDG template does not yet have a stud rod standard. Please contact the administrator.');
  return {kind, diameter, length};
}
// EPDM gasket identified by DN, PN and thickness in mm (3 unless written, e.g. DN400 PN10 5mm = 10).
function gasketItem(body, original) {
  let rest = ` ${body.replace(/\b(?:EPDM|gaskets?|full[-\s]*face|(?:with|wth)?\s*bolt\s*holes?)\b/gi, ' ')} `;
  if (/\b(?:bolt|nut|washer|stud|rod|SS316|SS16|HDG)\b/i.test(rest)) throw Error(`Only EPDM gaskets are accepted in a gasket order: ${original}. Send bolts, nuts, washers and stud rods as a separate SS316 or HDG order.`);
  const dn = rest.match(/\bDN\s*(\d+)\b/i), pn = rest.match(/\bPN\s*(\d+(?:\.\d+)?)\b/i);
  if (!dn || !pn) throw Error(`Gasket size not recognized: ${original}. Write the DN and PN, for example DN400 PN10 = 10.`);
  rest = rest.replace(dn[0], ' ').replace(pn[0], ' ');
  const thk = rest.match(/\b(\d+(?:\.\d+)?)\s*mm\b\s*(?:thk\.?|thick(?:ness)?)?/i);
  if (thk) rest = rest.replace(thk[0], ' ');
  if (/[^\s,;\/x-]/i.test(rest)) throw Error(`Gasket size not recognized: ${original}. Example: DN400 PN10 = 10, or DN400 PN10 5mm = 10 for another thickness.`);
  const thickness = thk ? Number(thk[1]) : 3;
  if (!(thickness > 0 && thickness <= 20) || !(Number(pn[1]) > 0)) throw Error(`Check the gasket thickness and PN: ${original}. Example: DN400 PN10 5mm = 10.`);
  return {kind: 'gasket', dn: positive(dn[1]), pn: Number(pn[1]), thickness};
}
function positive(value) {
  if (!/^(?:\d+|\d{1,3}(?:,\d{3})+)$/.test(value)) throw Error('Invalid quantity format. Use a whole number, for example 200 or 1,200.');
  const n = Number(value.replaceAll(',', ''));
  if (!Number.isSafeInteger(n) || n < 1 || n > 10000000) throw Error('Quantity must be a whole number between 1 and 10,000,000.');
  return n;
}
