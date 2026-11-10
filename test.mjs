import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseOrder,example,gasketExample} from './parser.mjs';
import {generate} from './excel.mjs';
import {latestDueWeek,loadOrders,nextWeek,orderNotice,summaryText,summaryWorkbook,weekOf} from './summary.mjs';
import {archiveDeleted,deleteBody,deleteSubject,findOrder,mailBody,mailSubject} from './mail.mjs';
import {formatDates,generateSteelForms,isSteelForm,parseSteelForm,steelExample,steelNotice} from './steel.mjs';
import {gasketOrderText,gridToOrderLines,isGasketTable,linesToGrid,orderTextFromTable,readGrid} from './table.mjs';
import {SpreadsheetFile,Workbook} from '@oai/artifact-tool';
const here=path.dirname(fileURLToPath(import.meta.url));
const ss=parseOrder(example);
assert.deepEqual(ss.sections[0].items.map(x=>x.qty),[2880,8640,5760,17280]);
assert.equal(parseOrder(example.replace('SS316','SS16')).material,'SS316');
const hdg=parseOrder(`C3D2 HDG
Title: WDL21 Grit Classifier Platform handrail
Sets: 50
M12x135mm HDG bolt = 1
M12 HDG nut = 3
M12x3mm HDG flat washer = 4
Title: WDL23 bioreactor piperack installation
Sets: 50
M16x100mm bolt = 1
M16 nut = 3
M16x3mm washer = 4`,'Test Requester');
assert.deepEqual(hdg.sections.map(s=>s.items.map(x=>x.qty)),[[50,150,200],[50,150,200]]);
const rod=parseOrder('C3D1 SS316\nTitle: WDL12 chemical support\nM16x2m SS stud rod = 20\nM16 washer = 200','Test');
assert.equal(rod.sections[0].items[0].length,2000);
assert.equal(rod.sections[0].items[1].length,3);
for (const invalid of [
  'C3D1 SS316\nTitle: WDL12\nM16 bolt = 200',
  'C3D1 SS316\nTitle: WDL12\nM16 HDG nut = 200',
  'C3D1 SS316\nTitle: WDL12\nM16 nut = 2,00',
  'C3D1 SS316\nTitle: WDL12\nM16 nut = -200',
  'C3D1 SS316\nTitle: WDL12\nM16 nut = 200\nadd 5%',
  'C3D1 SS316\nTitle: WDL12\nM16x65mm bolt',
  'C3D1 SS316\nTitle: WDL12\nM16 nut = 200\nTitle: empty'
]) assert.throws(()=>parseOrder(invalid,'Test'));
assert.equal(parseOrder('C3D1 SS316\nTitle: WDL12\nM16 bolt length 170 Qty 200','Test').sections[0].items[0].length,170);
// qty=70, qty: 70 and Quantity = 70 are accepted as the quantity delimiter (2026-10-07: a colleague's 2 m stud rod order written this way was refused).
const qtyEq=parseOrder('C3D1 SS316\nRequester Name: Chris Lim\nTitle: STUD ROD SUPPORT AND ANCHOR FOR FRP WDL12\nM16X2m stud rod qty=70\nM12X2m stud rod qty=70\nM10X2m stud rod qty=25\nM16 NUT qty=1000\nM12 NUT  qty=1000\nM16  FLAT WASHER qty=1000\nM12  FLAT WASHER qty=1000');
assert.equal(qtyEq.requester,'Chris Lim');
assert.deepEqual(qtyEq.sections[0].items.map(x=>[x.kind,x.diameter,x.length,x.qty]),[['rod',16,2000,70],['rod',12,2000,70],['rod',10,2000,25],['nut',16,null,1000],['nut',12,null,1000],['washer',16,3,1000],['washer',12,3,1000]]);
assert.equal(parseOrder('C3D1 SS316\nTitle: WDL12\nM16x2m stud rod qty: 70','Test').sections[0].items[0].qty,70);
assert.equal(parseOrder('C3D1 SS316\nTitle: WDL12\nM16 nut Quantity = 1,200','Test').sections[0].items[0].qty,1200);
assert.throws(()=>parseOrder('C3D1 SS316\nTitle: WDL12\nM16x2m stud rod qty=','Test'));
// Gaskets: DN, PN and thickness (3mm unless written) identify an item; matching gaskets combine within a title.
const gasket=parseOrder(`C3D1 Gasket
Title: WDL12 Pump Gallery 2 and 3 PSD
DN400 PN10 = 10
DN200 PN10 = 30
DN 200 PN 10 EPDM gasket = 20
DN200 PN16 = 5
DN150 PN10 5mm thk = 4
Title: WDL13 MBR Tank Header
Sets: 4
DN100 PN16 Full-Face with bolt holes = 2`,'Test Requester');
assert.equal(gasket.material,'EPDM');
assert.deepEqual(gasket.sections.map(s=>s.items.map(x=>[x.dn,x.pn,x.thickness,x.qty])),[[[400,10,3,10],[200,10,3,50],[200,16,3,5],[150,10,5,4]],[[100,16,3,8]]]);
assert.deepEqual(gasket.sections[0].items[1].parts,[30,20]);
assert.equal(gasket.sections[1].items[0].perSet,'4 sets x 2');
assert.equal(parseOrder(gasketExample).material,'EPDM');
for (const head of ['C3D2 EPDM','C3D2 EPDM Gasket','C3D2 gasket epdm','C3D2 Gaskets']) assert.equal(parseOrder(`${head}\nTitle: WDL21\nDN50 PN16 = 1`,'Test').material,'EPDM');
for (const invalid of [
  'C3D1 Gasket\nTitle: WDL12\nDN400 = 10',
  'C3D1 Gasket\nTitle: WDL12\nPN10 = 10',
  'C3D1 Gasket\nTitle: WDL12\nM16 nut = 10',
  'C3D1 Gasket\nTitle: WDL12\nDN400 PN10 SS316 = 10',
  'C3D1 Gasket\nTitle: WDL12\nDN400 PN10 30mm = 10',
  'C3D1 Gasket\nTitle: WDL12\nDN400 class 150 = 10',
  'C3D1 Gasket\nTitle: WDL12\nDN400 PN10',
  'C3D1 SS316\nTitle: WDL12\nDN200 gasket = 50',
  'C3D1 HDG\nTitle: WDL12\nDN200 PN10 = 50'
]) assert.throws(()=>parseOrder(invalid,'Test'));
// Puts items on the templates' old shared-formula rows (13-15, 30-32, 51-52) and the Total on row 56.
const long=parseOrder(`C3D1 SS316
Title: WDL12 Chemical Hanging Piperack Replacement
M16x2m stud rod = 20
M16x3mm washer = 200
M16 nut = 400
Title: WDL12 Cantilever Piperack Replacement
M20 nut = 120
M20x3mm washer = 60
M20x2m stud rod = 8
Title: WDL13 Bioreactor PR
M16 nut = 60
M20 nut = 1100
M30 nut = 120
M16x3mm washer = 50
M20x3mm washer = 1000
M30x4mm washer = 110
M30x2m stud rod = 22
Title: WDL13 MBR Filtrate Pump Monorail
M10 nut = 100
M12 nut = 400
M10x3mm washer = 65
M12x3mm washer = 285
M10x2m stud rod = 10
M12x2m stud rod = 10
Title: WDL13 DEOX RAS Crane
M12 nut = 70
M16 nut = 200
M20 nut = 180
M24 nut = 450
M12x3mm washer = 45
M16x3mm washer = 135
M20x3mm washer = 120
M24 washer = 300
Title: WDL22 Lamella Truss
M16x65mm bolt = 180
M20x65mm bolt = 860
M20x85mm bolt = 1300
M24x125mm bolt = 540
M30x125mm bolt = 330
M24 nut = 2080
M20 nut = 500
M24x4mm washer = 930
M20 nut = 60`,'Test Requester');
assert.deepEqual(long.sections.map(s=>s.items.map(x=>x.qty)),[[20,200,400],[120,60,8],[60,1100,120,50,1000,110,22],[100,400,65,285,10,10],[70,200,180,450,45,135,120,300],[180,860,1300,540,330,2080,560,930]]);
console.log('Parser checks passed: consolidation, lengths, units, sets, missing fields, material conflicts, gaskets.');
// Table orders (Excel file or photo): the grid becomes text order lines, then the normal parser consolidates.
const sampleGrid=[['SS316','Bolt(A4-70)','Nut (A4-80)','Washer'],['M24X210',630,630,1260],['M24X60( BOLT only)',378,'',378],['M24X50(Bolt only)',378,null,378],['M20X190',160,160,320],['M20X170',488,488,976],['M20X160',368,368,736],['M20x60(Bolt only )',340,'',''],['M20X50 (Bolt only)',410,'',''],['M20x40(Bolt only)',256,'',''],['M20x35(Bolt only)',96,'',''],['M16X100',64,64,128],['M16x35(Bolt only)',112,'',''],['Total','3,580','','']];
const tableLines=gridToOrderLines(sampleGrid);
assert.equal(tableLines.material,'SS316');
assert.deepEqual(tableLines.lines.slice(0,5),['M24x210mm bolt = 630','M24 nut = 630','M24 washer = 1260','M24x60mm bolt = 378','M24 washer = 378']);
assert.equal(tableLines.lines.length,24);
const tableText=orderTextFromTable('C3D1 SS16\nTitle: Knife Gate Valves',tableLines);
const tableOrder=parseOrder(tableText,'Test Requester');
assert.deepEqual(tableOrder.sections[0].items.map(x=>[x.kind,x.diameter,x.length,x.qty]),[['bolt',24,210,630],['nut',24,null,630],['washer',24,4,2016],['bolt',24,60,378],['bolt',24,50,378],['bolt',20,190,160],['nut',20,null,1016],['washer',20,3,2032],['bolt',20,170,488],['bolt',20,160,368],['bolt',20,60,340],['bolt',20,50,410],['bolt',20,40,256],['bolt',20,35,96],['bolt',16,100,64],['nut',16,null,64],['washer',16,3,128],['bolt',16,35,112]]);
assert.throws(()=>orderTextFromTable('C3D1 HDG\nTitle: X',tableLines),/table says SS316/);
assert.throws(()=>orderTextFromTable('C3D1 SS16',tableLines),/Add Title/);
assert.throws(()=>orderTextFromTable('Knife Gate Valves',tableLines),/first line of the caption/);
assert.throws(()=>gridToOrderLines([['Size','Bolt'],['M24',10]]),/needs a length/);
assert.throws(()=>gridToOrderLines([['Size','Bolt'],['M24X210','ten']]),/Quantity not recognized/);
assert.deepEqual(gridToOrderLines([['','Washer (4mm)','Stud rod'],['M16X2m','200','20']]).lines,['M16x4mm washer = 200','M16x2m stud rod = 20']);
// List tables (added 2026-10-02): one worded item per row, a Total Qty column, project and title in the top row.
// Rows as typed in Excel, and as OCR returns them from a photo (merged words, a Chinese translation in brackets).
const listGrid=[['C3D2','C3D2 WDL23 Cable Rack Type 1-4','5%'],['SN','Item Description','Total Qty incl. Backup (pcs)'],['1','HEXAGON BOLT: M10  Gr 8.8, L=50mm','2621'],['3','HEXAGONBOLT:M12 Gr8.8, L=45mm （六角螺栓：M12不锈钢8.8级， 长=45mm)','3169'],['','HEXAGONBOLT:M10 0Gr8.8,L=55mm','21'],['13','NUT: M10','2621'],['14','NUT:M12 （螺母：M12)','5,263'],['17','WASHER: M10 X 3mm THKS','5242'],['18','WASHER:M12X3mmTHKS （垫圈M12*3mm厚）','10526'],['','','']];
const listTable=gridToOrderLines(listGrid);
assert.deepEqual([listTable.project,listTable.material,listTable.title,listTable.percent],['C3D2','HDG','WDL23 Cable Rack Type 1-4',true]);
assert.deepEqual(listTable.lines,['M10x50mm bolt = 2621','M12x45mm bolt = 3169','M10x55mm bolt = 21','M10 nut = 2621','M12 nut = 5263','M10x3mm washer = 5242','M12x3mm washer = 10526']);
assert.equal(gridToOrderLines([['C3D2','C3D2WDL23CableRackType1-4 5%'],['SN','ItemDescription','TotalQtyincl.Backup(pcs)'],['1','NUT:M16','2428']]).title,'WDL23 Cable Rack Type 1-4');
const listText=orderTextFromTable('',listTable);
assert.equal(listText.split('\n').slice(0,2).join('|'),'C3D2 HDG|Title: WDL23 Cable Rack Type 1-4');
assert.equal(orderTextFromTable('C3D2 HDG / Title: Cable Rack',listTable).split('\n').slice(0,2).join('|'),'C3D2 HDG|Title: Cable Rack');
assert.equal(orderTextFromTable('HDG\nRequester Name: Someone',listTable).split('\n').slice(0,3).join('|'),'C3D2 HDG|Requester Name: Someone|Title: WDL23 Cable Rack Type 1-4');
const listOrder=parseOrder(listText,'Test Requester');
assert.deepEqual(listOrder.sections[0].items.map(x=>[x.kind,x.diameter,x.length,x.qty]),[['bolt',10,50,2621],['bolt',12,45,3169],['bolt',10,55,21],['nut',10,null,2621],['nut',12,null,5263],['washer',10,3,5242],['washer',12,3,10526]]);
assert.throws(()=>orderTextFromTable('C3D1 HDG',listTable),/table says C3D2/);
assert.throws(()=>orderTextFromTable('C3D2 SS16',listTable),/table says HDG/);
assert.throws(()=>gridToOrderLines([['Item','Qty'],['HEXAGON BOLT: M12 Gr 10.9, L=45mm','5']]),/Grade 10.9/);
assert.throws(()=>gridToOrderLines([['Item','Qty'],['HEXAGON BOLT: M12 Gr 8.8','5']]),/needs a length/);
assert.throws(()=>gridToOrderLines([['Item','Qty'],['NUT: M12','']]),/No quantity/);
assert.throws(()=>gridToOrderLines([['Item','Qty'],['BOLT & NUT: M12x45','5']]),/more than one item/);
assert.throws(()=>orderTextFromTable('',gridToOrderLines([['Item','Qty'],['NUT: M12','5']])),/project and material/);
// Numbered lists with size-first descriptions (added 2026-11-06, as OCR returned a colleague's photo): heading merged
// into one cell, a "Bolt, Nut & Washer A4-70" section label, washer thickness as "4 thk", the last row merged into one cell.
const numbered=gridToOrderLines([['Item Description Qty Unit','','',''],['','Bolt, Nut & Washer A4-70','',''],['17','M30 x 190mm Hexagon Bolt','48','Nos'],['17a','Hex Nut M30','48','Nos'],['17b','Flat Washer M30 (31 ID x 56 0D x 4 thk)','96','Nos'],['18a','HexNut M27','10','Nos'],['25b Flat Washer M24 (25 ID x 44 OD × 4 thk) 20 Nos','','','']]);
assert.equal(numbered.material,'SS316');
assert.deepEqual(numbered.lines,['M30x190mm bolt = 48','M30 nut = 48','M30x4mm washer = 96','M27 nut = 10','M24x4mm washer = 20']);
assert.equal(parseOrder(orderTextFromTable('C3D1 SS316\nTITLE: WDL13 Valves\nRequestor: Ben',numbered),'Sender').requester,'Ben');
assert.throws(()=>orderTextFromTable('C3D1 HDG\nTitle: X',numbered),/table says SS316/);
// Photo of a Size / Grade / Qty table (added 2026-10-09): the item word may follow the size, and the grade column gives the material.
const photoGrid=[['SL.no','Size','Grade','Qty','Remarks'],['','Bolt M16x65mm Full Thread Long','SS A4-70','1650','Hex.Bolt'],['','M16 Nut','SS A4-80','1650',''],['','M16Nut','SSA4-80','512',''],['','M16 Washer','SS316L','3300',''],['','Bolt M27x 110mm Full Thread Long','SS A4-70','42','Hex.Bolt']];
assert.deepEqual(gridToOrderLines(photoGrid).lines,['M16x65mm bolt = 1650','M16 nut = 1650','M16 nut = 512','M16 washer = 3300','M27x110mm bolt = 42']);
assert.equal(gridToOrderLines(photoGrid).material,'SS316');
assert.throws(()=>orderTextFromTable('C3D2 HDG\nTitle: X',gridToOrderLines(photoGrid)),/table says SS316/);
assert.throws(()=>gridToOrderLines([...photoGrid,['','M36 Nut','SS A4-80','','']]),/No quantity/);
// Description / Size / Length / Qty columns (2026-10-09), as OCR returns them (words run together, title unspaced).
const sizeGrid=[['MATERIALLISTFORBIOSORPTIONBLOWERS(EVERSTAR)','','','',''],['DESCRIPTION','SIZE','SCHEDULE','LENGTH (MM)','Qty'],['BS4190BOLTA4-70','M27','','145','40'],['BS4190NUTSA4-70','M27','','','160'],['BS4190FLASHWASHERA4-70','M27','','','320'],['BS 4190 BOLT A4-70','M20','','90','240'],['BS 4190 FLAT WASHER A4-70','M20','','','680'],['S.NO','','','',''],['10','','','','']];
const sizeTable=gridToOrderLines(sizeGrid);
assert.deepEqual([sizeTable.lines,sizeTable.material,sizeTable.title],[['M27x145mm bolt = 40','M27 nut = 160','M27 washer = 320','M20x90mm bolt = 240','M20 washer = 680'],'SS316',null]);
assert.throws(()=>orderTextFromTable('C3D1 SS16',sizeTable),/Add Title/);
assert.throws(()=>gridToOrderLines([sizeGrid[1],['BS 4190 BOLT A4-70','M20','','','24']]),/needs a length/);
// Washer thickness (2026-10-09): M20 and below 3 mm, M24 and above 4 mm by default and never thinner.
assert.deepEqual(parseOrder('C3D2 SS316\nTitle: X\nM8 washer = 1\nM20 washer = 1\nM24 washer = 1\nM33 washer = 1\nM46 washer = 1\nM36x5mm washer = 1','T').sections[0].items.map(x=>[x.diameter,x.length]),[[8,3],[20,3],[24,4],[33,4],[46,4],[36,5]]);
assert.throws(()=>parseOrder('C3D2 SS316\nTitle: X\nM24x3mm washer = 1','T'),/at least 4mm/);
// Gasket lists (2026-10-09): typed or as a caption, or from a photo / Excel grid.
const gasketList='C3D2 Fibre Reinforced EPDM 3mm Thk Gasket for WDL21 Dampers\nBS EN1092-01\nSl.no DN Size Pressure Rating Qty\n1 DN100 PN10 106\n2 DN150 PN2.5 102\n3 DN150 PN10 304\n15 DN1800 PN10 2';
assert.ok(isGasketTable(gasketList) && !isGasketTable(gasketExample) && !isGasketTable(example));
assert.equal(gasketOrderText(gasketList),'C3D2 Gasket\nTitle: WDL21 Dampers\nDN100 PN10 3mm = 106\nDN150 PN2.5 3mm = 102\nDN150 PN10 3mm = 304\nDN1800 PN10 3mm = 2');
const gasketOrder=parseOrder(gasketOrderText(gasketList),'T');
assert.deepEqual(gasketOrder.sections[0].items.map(x=>[x.dn,x.pn,x.thickness,x.qty]),[[100,10,3,106],[150,2.5,3,102],[150,10,3,304],[1800,10,3,2]]);
assert.equal(gasketOrderText('C3D1 EPDM Gasket\nTitle: Pump Gallery',[['Sl.no','DN Size','Pressure Rating','Qty'],['1','DN 400','PN10','10'],['2','DN200','PN 16','1,200']]),'C3D1 Gasket\nTitle: Pump Gallery\nDN400 PN10 = 10\nDN200 PN16 = 1200');
assert.equal(gasketOrderText('C3D2 Gasket\nRequester: Dana\nTitle: WDL21 Dampers',[['St.no','DN Size','Pressure Rating','Qty'],['1','1|DN100','PN10','106'],['2','2DN150','PN2.5','102'],['','DN350','PN10','14'],['11','LDN800','PN10','4']]),'C3D2 Gasket\nRequester: Dana\nTitle: WDL21 Dampers\nDN100 PN10 = 106\nDN150 PN2.5 = 102\nDN350 PN10 = 14\nDN800 PN10 = 4');
// Qty and PN in their own columns, PN as a bare number, thickness from a "Gasket 3mm EPDM" row (added 2026-11-06).
const pnColumn=[['Item','Description','Qty','PN'],['','VALVES AND FLANGE ADAPTOR','',''],['','Gasket 3mm EPDM','',''],['12','DN800','4','10'],['14','DN600','16','10'],['16','DN400','4','PN16']];
assert.equal(gasketOrderText('C3D1 Gasket\nRequester: Ben\nTitle: Valves',pnColumn),'C3D1 Gasket\nRequester: Ben\nTitle: Valves\nDN800 PN10 3mm = 4\nDN600 PN10 3mm = 16\nDN400 PN16 3mm = 4');
assert.throws(()=>gasketOrderText('C3D1 Gasket\nTitle: X',[['Item','Description','Qty','PN'],['12','DN800','4','']]),/DN800 has no PN rating/);
// Positioned OCR lines of the colleague's photo (Qty figures re-read from close crops) become that grid.
const ocrLines={width:1000,lines:[[{text:'Item',x0:29,x1:93},{text:'Description',x0:338,x1:486},{text:'Qty',x0:749,x1:805},{text:'PN',x0:895,x1:936}],[{text:'VALVESANDFLANGEADAPTOR',x0:121,x1:519}],[{text:'Gasket3mmEPDM',x0:120,x1:362}],[{text:'12',x0:44,x1:77},{text:'DN800',x0:120,x1:204},{text:'4',x0:776,x1:778},{text:'10',x0:900,x1:933}],[{text:'14',x0:44,x1:77},{text:'DN600',x0:121,x1:203},{text:'16',x0:761,x1:794},{text:'10',x0:899,x1:933}]]};
assert.deepEqual(linesToGrid(ocrLines).slice(2),[['','Gasket3mmEPDM','',''],['12 DN800','','4','10'],['14 DN600','','16','10']]);
assert.equal(gasketOrderText('C3D1 Gasket\nTitle: Valves',linesToGrid(ocrLines)),'C3D1 Gasket\nTitle: Valves\nDN800 PN10 3mm = 4\nDN600 PN10 3mm = 16');
assert.equal(linesToGrid({lines:[[{text:'DN800',x0:1,x1:9}]]}),null);
assert.throws(()=>gasketOrderText('C3D1 Gasket\nTitle: X',[['1','DN400','Class 150','10']]),/no PN rating/);
assert.throws(()=>gasketOrderText('C3D1 Gasket\nTitle: X',[['1','DN400','PN10','']]),/No quantity/);
assert.throws(()=>gasketOrderText('C3D1 Neoprene Gasket for X\n1 DN100 PN10 5'),/Only EPDM/);
assert.throws(()=>gasketOrderText('C3D1 EPDM Gasket\n1 DN100 PN10 5'),/Add Title/);
const gridBook=Workbook.create(); const gridSheet=gridBook.worksheets.add('Sheet1'); gridSheet.getRange('A1:D3').values=[['SS316','Bolt','Nut','Washer'],['M24X210',630,630,1260],['M20x35(Bolt only)',96,null,null]];
const gridFile=path.join(os.tmpdir(),`grid-${Date.now()}.xlsx`); await (await SpreadsheetFile.exportXlsx(gridBook)).save(gridFile);
assert.deepEqual(await readGrid(gridFile),[['SS316','Bolt','Nut','Washer'],['M24X210','630','630','1260'],['M20x35(Bolt only)','96','','']]);
await fs.rm(gridFile,{force:true});
console.log('Table order checks passed: grid to lines, caption checks, consolidation, Excel reading.');
for (const [name,order] of [['ss316-test',ss],['hdg-test',hdg],['long-test',long],['gasket-test',gasket],['table-test',tableOrder],['list-table-test',listOrder]]) console.log(await generate(order,name,true));
console.log('All workbook exports and numeric/formula verification passed.');
// Weekly summaries: Monday-to-Sunday weeks in Singapore time, each due the following Monday at 09:00.
assert.equal(weekOf(Date.parse('2026-09-20T23:59:00+08:00')),'2026-09-14');
assert.equal(weekOf(Date.parse('2026-09-21T00:00:00+08:00')),'2026-09-21');
assert.equal(latestDueWeek(Date.parse('2026-09-21T08:59:00+08:00')),'2026-09-07');
assert.equal(latestDueWeek(Date.parse('2026-09-21T09:00:00+08:00')),'2026-09-14');
assert.equal(nextWeek('2026-12-28'),'2027-01-04');
const outputs=await fs.mkdtemp(path.join(os.tmpdir(),'fastener-summary-'));
for (const [name,receivedAt,order,status] of [
  ['hdg','2026-09-15T02:00:00Z',hdg,'sent'],
  ['ss','2026-09-14T01:00:00Z',ss,'uncertain'],
  ['gasket','2026-09-17T02:00:00Z',gasket,'sent'],
  ['previous-week','2026-09-13T15:59:00Z',ss,'sent'],
  ['generation-failed','2026-09-16T02:00:00Z',ss,null],
]) {
  await fs.mkdir(path.join(outputs,name));
  await fs.writeFile(path.join(outputs,name,'request.json'),JSON.stringify({receivedAt,order}));
  if (status) await fs.writeFile(path.join(outputs,name,'delivery.json'),JSON.stringify({status}));
}
const week=await loadOrders(outputs,'2026-09-14');
await fs.rm(outputs,{recursive:true});
assert.deepEqual(week.map(o=>[o.project,o.material,o.delivered]),[['C3D1','SS316',false],['C3D2','HDG',true],['C3D1','EPDM',true]]);
const text=summaryText(week,'2026-09-14');
assert.ok(text.startsWith('Weekly order summary\nMon 14 Sep – Sun 20 Sep 2026 · 3 orders\n\n1. C3D1 SS316 · Mon 14 Sep · Nicholas (delivery unconfirmed)\nWDL13 Membrane Tank Pipe Header Valve Connection\n• FULL THREADED HEXAGON BOLT: M16, L=110mm × 2,880\n• HEXAGON NUT: M16 × 8,640\n'),text);
assert.ok(text.includes('\n\n2. C3D2 HDG · Tue 15 Sep · Test Requester\nWDL21 Grit Classifier Platform handrail\n• FULL THREADED HEXAGON BOLT: M12, L=135mm × 50\n'),text);
assert.ok(text.includes('\n\n3. C3D1 EPDM · Thu 17 Sep · Test Requester\nWDL12 Pump Gallery 2 and 3 PSD\n• SIZE: DN400 3MM THK Full-Face with Bolt Holes, PN10 × 10\n• SIZE: DN200 3MM THK Full-Face with Bolt Holes, PN10 × 50\n'),text);
assert.equal(summaryText([],'2026-09-14',true),'Weekly order summary\nMon 14 Sep – Sun 20 Sep 2026 · no orders so far');
const busy=summaryText(Array(40).fill(week[0]),'2026-09-14');
assert.ok(busy.length<=4096 && busy.endsWith('\n…the rest is in the attached Excel file.'),busy.slice(-200));
console.log(await summaryWorkbook(week,'2026-09-14',path.join(path.dirname(fileURLToPath(import.meta.url)),'outputs','summary-test')));
console.log('Weekly summary checks passed: week boundaries, order selection, message text and length, Excel list.');
// Administrator copy of a colleague's order: caption names the sender, order and items within Telegram's 1024-character caption limit.
const notice=orderNotice(hdg,'Alex Tan (@alextan)',true);
assert.equal(notice,'New order from Alex Tan (@alextan)\nC3D2 HDG · Requester: Test Requester\nWDL21 Grit Classifier Platform handrail\n• FULL THREADED HEXAGON BOLT: M12, L=135mm × 50\n• HEXAGON NUT: M12 × 150\n• FLAT WASHER: M12 X 3mm thk. × 200\nWDL23 bioreactor piperack installation\n• FULL THREADED HEXAGON BOLT: M16, L=100mm × 50\n• HEXAGON NUT: M16 × 150\n• FLAT WASHER: M16 X 3mm thk. × 200');
assert.ok(orderNotice(gasket,'Alex Tan',false).startsWith('New order from Alex Tan (delivery to the sender unconfirmed)\nC3D1 EPDM · Requester: Test Requester\nWDL12 Pump Gallery 2 and 3 PSD\n• SIZE: DN400 3MM THK Full-Face with Bolt Holes, PN10 × 10\n'));
const bigNotice=orderNotice(long,'Alex Tan',true);
assert.ok(bigNotice.length<=1024 && bigNotice.endsWith('\n…see the Excel file for the full list.'),bigNotice.slice(-200));
console.log('Administrator notice checks passed: caption text, unconfirmed delivery note, caption length.');
// /sendemail: subject and body wording, and which order a chat's /sendemail refers to.
assert.equal(mailSubject(hdg),'C3D2 HDG Fasteners · Test Requester · WDL21 Grit Classifier Platform handrail; WDL23 bioreactor piperack installation');
assert.equal(mailSubject(gasket),'C3D1 EPDM Gaskets · Test Requester · WDL12 Pump Gallery 2 and 3 PSD; WDL13 MBR Tank Header');
const longSubject=mailSubject({...ss,sections:[{title:'W'.repeat(300),items:[]}]});
assert.ok(longSubject.length===255 && longSubject.endsWith('…'),longSubject.length);
assert.equal(mailBody(ss,'Alex Tan','C:\\x\\C3D1 SS316 Fastener Order.xlsx',new Date('2026-09-15T06:05:00Z')),'WDL13 Membrane Tank Pipe Header Valve Connection\n• FULL THREADED HEXAGON BOLT: M16, L=110mm × 2,880\n• HEXAGON NUT: M16 × 8,640\n• FULL THREADED HEXAGON BOLT: M16, L=130mm × 5,760\n• FLAT WASHER: M16 X 3mm thk. × 17,280\n\nRequester: Nicholas · Sent by: Alex Tan (Telegram) · 15 Sep 2026 14:05\nAttachment: C3D1 SS316 Fastener Order.xlsx');
const archive=await fs.mkdtemp(path.join(os.tmpdir(),'fastener-mail-'));
for (const [name,order,delivery,withFile] of [
  ['100_5',hdg,{status:'sent',messageId:6,adminMessageId:91},true],
  ['100_9',ss,{status:'sent',messageId:10,adminMessageId:95},true],
  ['100_12',gasket,null,false],
  ['200_3',gasket,{status:'sent',messageId:4},true],
]) {
  await fs.mkdir(path.join(archive,name));
  await fs.writeFile(path.join(archive,name,'request.json'),JSON.stringify({order}));
  if (delivery) await fs.writeFile(path.join(archive,name,'delivery.json'),JSON.stringify(delivery));
  if (withFile) await fs.writeFile(path.join(archive,name,`${order.project} ${order.material} Order.xlsx`),'');
}
const latest=await findOrder(archive,{chat:100,owner:900});
assert.equal(latest.order.material,'SS316');
assert.ok(latest.file.endsWith('100_9\\C3D1 SS316 Order.xlsx') || latest.file.endsWith('100_9/C3D1 SS316 Order.xlsx'),latest.file);
assert.equal((await findOrder(archive,{chat:100,owner:900,replyTo:6})).order.material,'HDG');
assert.equal(await findOrder(archive,{chat:100,owner:900,replyTo:91}),null);
assert.equal((await findOrder(archive,{chat:900,owner:900,replyTo:91})).order.material,'HDG');
assert.equal(await findOrder(archive,{chat:900,owner:900}),null);
assert.equal((await findOrder(archive,{chat:200,owner:900})).order.material,'EPDM');
// /delete: notice wording, and the folder moves out of the records (lookup, summary) into deleted/.
assert.equal(deleteSubject(hdg),`DELETED · ${mailSubject(hdg)}`);
assert.ok(deleteSubject({...ss,sections:[{title:'W'.repeat(300),items:[]}]}).length===255);
const deleteText=deleteBody(ss,'Alex Tan','C:\\x\\C3D1 SS316 Fastener Order.xlsx',{when:new Date('2026-10-12T06:05:00Z')});
assert.ok(deleteText.startsWith('Alex Tan (Telegram) deleted this order file on 12 Oct 2026 14:05. It is removed from the order records'),deleteText);
assert.ok(!deleteText.includes('already emailed') && deleteText.endsWith('Requester: Nicholas\nDeleted file attached for your records: C3D1 SS316 Fastener Order.xlsx'),deleteText.slice(-200));
assert.ok(deleteBody(ss,'Alex Tan','x.xlsx',{when:new Date('2026-10-12T06:05:00Z'),emailedAt:'2026-10-12T05:00:00Z'}).includes('\n\nThis order was already emailed to you on 12 Oct 2026 13:00. Please disregard that email.\n\n'));
const toDelete=await findOrder(archive,{chat:100,owner:900});
assert.equal(toDelete.order.material,'SS316');
const movedTo=await archiveDeleted(toDelete.dir,path.join(archive,'..',`deleted-${path.basename(archive)}`),{at:'2026-10-12T06:05:00Z',sender:'Alex Tan'});
assert.ok((await fs.readdir(movedTo)).includes('deleted.json') && (await fs.readdir(movedTo)).includes('C3D1 SS316 Order.xlsx'));
assert.equal((await findOrder(archive,{chat:100,owner:900})).order.material,'HDG');
assert.equal(await findOrder(archive,{chat:100,owner:900,replyTo:10}),null);
await fs.rm(path.dirname(movedTo),{recursive:true});
await fs.rm(archive,{recursive:true});
console.log('Email checks passed: subject, body, order lookup by reply and by latest order, delete notice and move.');
// Steel-work Word forms: intake detection, parsing, the updater's date spellings, and a real generation run on copies of the templates.
assert.ok(isSteelForm(steelExample) && isSteelForm('C3D2 Steel Form\nRef: 049') && isSteelForm('Please update which project forms as following:\n\nProject Forms: C3D1'));
assert.ok(!isSteelForm(example) && !isSteelForm(gasketExample) && !isSteelForm('C3D1 HDG\nTitle: WDL12\nM16 nut = 200'));
const steel=parseSteelForm(steelExample,new Date('2026-09-21T01:00:00Z'));
assert.deepEqual(steel,{project:'C3D1',ref:'UESH/C3D1/PM/STPR/048',nnn:'048',currentDate:'21-September-2026',descLines:['WDL12 - FA DUCTS PST TANK 1 TO 6 ( GROUND SUPPORTS ) AT +15.800 SHD','WDL15 - ROP FLUSH TANK AREA AT +14.600 SHD'],batch:'1',tonnage:'5',delivery:'2/10/26'});
assert.equal(parseSteelForm(steelExample,new Date('2026-09-20T16:30:00Z')).currentDate,'21-September-2026'); // today is the Singapore date
const steel2=parseSteelForm('C3D2 Steel Form\nRef: 49\nDate: 18/8/26\nDescription of Steel Work:\nWDL22- EICA Cable Racks Types\nBatch: 5\nTonnage: 100\nDelivery Date: Type3-20/9/26');
assert.deepEqual([steel2.project,steel2.ref,steel2.currentDate,steel2.descLines,steel2.batch,steel2.tonnage,steel2.delivery],['C3D2','UESH/C3D2/PM/STPR/049','18-August-2026',['WDL22- EICA Cable Racks Types'],'5','100','Type3-20/9/26']);
assert.equal(parseSteelForm(steelExample.replace('Current Date: today','Current Date: 4 September 2026')).currentDate,'04-September-2026');
assert.throws(()=>parseSteelForm(steelExample.replace('Project Forms: C3D1','Project Forms: C3D1 / C3D2')),/Choose one project/);
assert.throws(()=>parseSteelForm(steelExample.replace('Estimated Tonnage: 5\n','')),/Missing: Estimated Tonnage/);
assert.throws(()=>parseSteelForm(steelExample.replace('UESH/C3D1/PM/STPR/048','UESH/C3D2/PM/STPR/048')),/belongs to C3D2/);
assert.throws(()=>parseSteelForm(steelExample.replace('2/10/26','31/2/26')),/not a real date/);
assert.throws(()=>parseSteelForm(steelExample.replace('2/10/26','next week')),/Delivery Date not recognized/);
assert.throws(()=>parseSteelForm(steelExample+'\nRemarks: urgent'),/Line not recognized/);
assert.deepEqual(formatDates('2/10/26'),{letter:'2-October-2026',request:'2 October 2026',long:'2nd October 2026',short:'2/10/26'});
assert.equal(formatDates('04-September-2026').letter,'4-September-2026');
assert.equal(formatDates('11/9/26').long,'11th September 2026');
assert.equal(formatDates('23/9/26').long,'23rd September 2026');
assert.equal(formatDates('Type3-20/9/26').long,'Type3-20/9/26');
const steelCaption=steelNotice(steel,{outDir:'X\\C3D1_STPR_048_Output_rev2',rev:2,layout:{ok:true,pages:{written:1,request:1}}});
assert.ok(steelCaption.startsWith('C3D1 · UESH/C3D1/PM/STPR/048 · Batch 1 · 5 t\nDate 21-September-2026 · Delivery 2nd October 2026\nWDL12 - FA DUCTS') && steelCaption.includes('Word layout OK: 1 page letter, 1 page request form') && steelCaption.endsWith('Saved: C3D1_STPR_048_Output_rev2 (revision 2; the earlier submission with this reference was kept)'),steelCaption);
assert.ok(steelNotice(steel,{outDir:'X\\C3D1_STPR_048_Output',rev:1,layout:{ok:false,unavailable:true,problems:['x']}}).includes('layout not visually checked (Word unavailable)'));
console.log('Steel form checks passed: detection, parsing, dates, captions.');
let steelCfg=null;
try {steelCfg=JSON.parse(await fs.readFile(path.join(here,'steel.json'),'utf8'));} catch(e) {if(e.code!=='ENOENT') throw e;}
if (steelCfg?.root) {
  // Generate into a temporary copy of the handoff so the canonical outputs stay untouched.
  const tmpRoot=await fs.mkdtemp(path.join(os.tmpdir(),'steel-forms-'));
  await fs.mkdir(path.join(tmpRoot,'outputs'));
  await fs.copyFile(path.join(steelCfg.root,'outputs','update_steel_forms.py'),path.join(tmpRoot,'outputs','update_steel_forms.py'));
  for (const dir of ['C3D1 Forms','C3D2 Forms']) await fs.cp(path.join(steelCfg.root,'outputs',dir),path.join(tmpRoot,'outputs',dir),{recursive:true});
  const cfg={root:tmpRoot,python:steelCfg.python}, render=!process.env.STEEL_SKIP_RENDER;
  const r1=await generateSteelForms(steel,cfg,{render});
  assert.ok(r1.verify.ok,JSON.stringify(r1.verify.problems));
  assert.deepEqual(r1.verify.signatures.written,['UES HOLDINGS PTE LTD','Kenneth','Project Manager']);
  assert.equal(r1.verify.signatures.request.length,4);
  assert.equal(r1.rev,1); assert.ok(r1.outDir.endsWith('C3D1_STPR_048_Output'));
  assert.deepEqual(JSON.parse(await fs.readFile(r1.jsonFile,'utf8')),{our_reference_number:'UESH/C3D1/PM/STPR/048',current_date:'21-September-2026',description_of_steel_work:steel.descLines,batch_number:'1',estimated_tonnage:'5',required_steel_delivery_date:'2/10/26'});
  if (render) {assert.ok(r1.layout.ok,JSON.stringify(r1.layout)); for (const pdf of Object.values(r1.layout.pdfs)) await fs.access(pdf);}
  const r2=await generateSteelForms(steel,cfg,{render:false});
  assert.equal(r2.rev,2); assert.ok(r2.outDir.endsWith('C3D1_STPR_048_Output_rev2') && r2.jsonFile.endsWith('submission_details_C3D1_STPR_048_rev2.json'));
  const r3=await generateSteelForms(steel2,cfg,{render});
  assert.ok(r3.verify.ok,JSON.stringify(r3.verify.problems));
  assert.equal(r3.verify.signatures.request[1],'Kenneth Yeo, PROJECT MANAGER');
  if (render) assert.ok(r3.layout.ok,JSON.stringify(r3.layout));
  await fs.rm(tmpRoot,{recursive:true});
  console.log(`Steel form generation passed: C3D1 and C3D2 forms verified${render?' and rendered with Word':''}; repeat reference saved as revision 2.`);
} else console.log('Steel form generation skipped: steel.json not found.');
process.exit(0);
