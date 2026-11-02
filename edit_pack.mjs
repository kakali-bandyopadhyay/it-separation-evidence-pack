import fs from 'node:fs/promises';
import { pathToFileURL } from 'node:url';
import { FileBlob, SpreadsheetFile } from 'file:///C:/Users/KBand/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const root = 'C:/Projects/it-separation-evidence-pack';
const xlsx = `${root}/it_separation_evidence_pack.xlsx`;
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(xlsx));
function parseCsv(text) {
  const rows=[]; let row=[], cell='', quoted=false;
  for(let i=0;i<text.length;i++) { const ch=text[i];
    if(quoted) { if(ch==='"'&&text[i+1]==='"'){cell+='"';i++;} else if(ch==='"') quoted=false; else cell+=ch; }
    else if(ch==='"') quoted=true;
    else if(ch===','){row.push(cell);cell='';}
    else if(ch==='\n'){row.push(cell.replace(/\r$/,''));rows.push(row);row=[];cell='';}
    else cell+=ch;
  }
  if(cell.length||row.length){row.push(cell.replace(/\r$/,''));rows.push(row);}
  return rows.filter(r=>r.some(v=>v!==''));
}
const app = wb.worksheets.getItem('App_Disposition');
const cost = wb.worksheets.getItem('Cost_Synergy');
const checks = wb.worksheets.getItem('Checks');
const dash = wb.worksheets.getItem('Dashboard');
const tsa = wb.worksheets.getItem('TSA_Schedule');

// Restore the original CSV waves and original disposition rules.
const appCsv = parseCsv(await fs.readFile(`${root}/data/apps.csv`,'utf8'));
const waveIndex = appCsv[0].indexOf('wave');
const sourceWaves = appCsv.slice(1).map(row=>[row[waveIndex]]);
app.getRange('N5:N64').values = sourceWaves;
app.getRange('U4:U64').clear({applyTo:'all'});
app.getRange('T4:T4').values = [['Final disposition']];
app.getRange('T5:T64').formulas = Array.from({length:60}, (_,i) => {
  const r = i + 5;
  return [`=IF(P${r}<>"",P${r},O${r})`];
});
app.getRange('O5:O64').formulas = Array.from({length:60}, (_,i) => {
  const r=i+5;
  return [`=IF(K${r}="Y","TSA-then-replace",IF(AND(H${r}="Y",I${r}="Low"),"Transfer",IF(H${r}="Consent","Clone",IF(I${r}="High","Replicate","Replace"))))`];
});
app.getRange('P54:Q54').clear({applyTo:'contents'});
app.getRange('P58:Q58').clear({applyTo:'contents'});
app.getRange('P14:Q14').values = [['Retire','Synthetic example: function absorbed by buyer platform']];
app.getRange('P5:P64').dataValidation = {
  allowBlank: true,
  list: { inCellDropDown: true, source: ['Transfer','Clone','Replicate','Replace','Retire','TSA-then-replace'] },
};
app.getRange('R5:R64').formulas = Array.from({length:60}, (_,i) => {
  const r=i+5;
  return [`=IF(T${r}="Retire","N",IF(COUNTIF(TSA_Schedule!$K$45:$K$97,A${r})>0,"Y","N"))`];
});
app.getRange('B69:C74').formulas = Array.from({length:6}, (_,i) => {
  const r=i+69;
  return [`=COUNTIF($T$5:$T$64,A${r})`,`=SUMIF($T$5:$T$64,A${r},$L$5:$L$64)`];
});
app.getRange('B77').formulas = [['=COUNTIFS($J$5:$J$64,1,$T$5:$T$64,"")']];
app.getRange('T4').format = app.getRange('O4').format;
app.getRange('T5:T64').format = app.getRange('O5:O64').format;
app.getRange('T:T').format.columnWidth = 21;
app.getRange('P14:Q14').format.font = {name:'Calibri',size:9,color:'#0070C0'};
app.getRange('P14:Q14').format.wrapText = true;
app.getRange('Q:Q').format.columnWidth = 42;
app.getRange('14:14').format.rowHeight = 30;

// Two period-consistent bridges replace the previous mixed-period net reference.
cost.getRange('H16:I33').clear({applyTo:'contents'});
cost.getRange('H25:U27').unmerge();
cost.getRange('H32:I32').unmerge();
cost.getRange('H25:U27').clear({applyTo:'formats'});
cost.getRange('H16:I17').values = [['Bridge A | Steady-state annual IT run-rate (synthetic)',null],['Component','GBP / year']];
cost.getRange('H18:H22').values = [['Allocated baseline'],['+ Dis-synergy'],['Standalone run-rate'],['- Identified synergies (midpoint, annual, steady state)'],['Target run-rate']];
cost.getRange('I18:I22').formulas = [['=B12'],['=D12'],['=SUM(I18:I19)'],['=SUM(Q6:Q20)'],['=I20-I21']];
cost.getRange('H24:I25').values = [['Bridge B | One-time cost to separate, Day 1 to TSA exit (synthetic)',null],['Component','GBP']];
cost.getRange('H26:H31').values = [['One-off Day 1 wave'],['One-off Wave 2'],['One-off Wave 3'],['Planned TSA fees (buyer side)'],['One-off separation cost'],['One-time cost to separate (one-off + TSA)']];
cost.getRange('I26:I31').formulas = [['=G7'],['=G6'],['=G8'],['=TSA_Schedule!R5'],['=SUM(I26:I28)'],['=SUM(I26:I29)']];
cost.getRange('H33:I33').merge();
cost.getRange('H33').values = [['TSA fees replace standalone run cost for covered services during the TSA period; they are not added to the run-rate.']];
cost.getRange('H16:I17').format = {fill:'#1F3864',font:{name:'Calibri',size:10,bold:true,color:'#FFFFFF'},wrapText:true,verticalAlignment:'center'};
cost.getRange('H24:I25').format = {fill:'#1F3864',font:{name:'Calibri',size:10,bold:true,color:'#FFFFFF'},wrapText:true,verticalAlignment:'center'};
cost.getRange('H18:H22').format = {font:{name:'Calibri',size:9,color:'#172B4D'},wrapText:true};
cost.getRange('H26:H31').format = {font:{name:'Calibri',size:9,color:'#172B4D'},wrapText:true};
cost.getRange('I18:I22').format.numberFormat = '#,##0';
cost.getRange('I26:I31').format.numberFormat = '#,##0';
cost.getRange('I22').format.font = {name:'Calibri',size:10,bold:true,color:'#000000'};
cost.getRange('I30:I31').format.font = {name:'Calibri',size:10,bold:true,color:'#000000'};
cost.getRange('H33').format = {font:{name:'Calibri',size:8,italic:true,color:'#667085'},wrapText:true,verticalAlignment:'top'};
cost.getRange('H:H').format.columnWidth = 31;
cost.getRange('I:I').format.columnWidth = 16;
cost.getRange('H16:I16').format.rowHeight = 30;
cost.getRange('H24:I24').format.rowHeight = 34;
cost.getRange('H33:I33').format.rowHeight = 40;

// Core checks: independent TSA arithmetic, four bridge ties, and override rationale completeness.
tsa.getRange('S4').values = [['Independent months (start/end dates)']];
tsa.getRange('S5:S39').formulas = Array.from({length:35},(_,i)=>{
  const r=i+5;
  return [`=(YEAR(F${r})-YEAR(E${r}))*12+MONTH(F${r})-MONTH(E${r})+1`];
});
tsa.getRange('S4:S39').copyFrom(tsa.getRange('G4:G39'),'all');
tsa.getRange('S4').values = [['Independent months (start/end dates)']];
tsa.getRange('S5:S39').formulas = Array.from({length:35},(_,i)=>{
  const r=i+5;
  return [`=(YEAR(F${r})-YEAR(E${r}))*12+MONTH(F${r})-MONTH(E${r})+1`];
});
tsa.getRange('S:S').format.columnWidth = 28;
tsa.getRange('S4').format.wrapText = true;
checks.getRange('A6:C6').values = [['TSA total equals independent date-based SUMPRODUCT',null,'Seller-to-buyer dates x cost basis x mark-up; avoids Months, Monthly Fee and line totals']];
checks.getRange('B6').formulas = [['=IF(ABS(TSA_Schedule!$R$5-SUMPRODUCT((TSA_Schedule!$C$5:$C$39="Seller to buyer")*TSA_Schedule!$S$5:$S$39*TSA_Schedule!$I$5:$I$39*(1+README_Assumptions!$B$14)))<0.01,"PASS","FAIL")']];
checks.getRange('A10:C10').values = [['Independent ties: both bridges',null,'Bridge B one-off and TSA; Bridge A tower baseline and synergy list midpoint']];
checks.getRange('B10').formulas = [['=IF(AND(ABS(Cost_Synergy!$I$30-SUM(App_Disposition!$L$5:$L$64))<0.01,ABS(Cost_Synergy!$I$29-TSA_Schedule!$G$44)<0.01,ABS(Cost_Synergy!$I$18-SUM(Cost_Synergy!$B$6:$B$11))<0.01,ABS(Cost_Synergy!$I$21-SUM(Cost_Synergy!$Q$6:$Q$20))<0.01),"PASS","FAIL")']];
checks.getRange('A13:C13').values = [['Every override has a rationale',null,'Each nonblank override has a nonblank rationale']];
checks.getRange('B13').formulas = [['=IF(SUMPRODUCT(--(App_Disposition!$P$5:$P$64<>""),--(App_Disposition!$Q$5:$Q$64=""))=0,"PASS","FAIL")']];
checks.getRange('A13:C13').format = checks.getRange('A12:C12').format;
checks.getRange('C174:D189').format.numberFormat = 'yyyy-mm-dd';

// Dashboard headline metrics and a two-bridge chart data block.
dash.getRange('D5:E5').values = [['One-off separation cost GBP',null]];
dash.getRange('E5').formulas = [['=Cost_Synergy!I30']];
dash.getRange('E5').format.numberFormat = '#,##0';
dash.getRange('D9:E10').values = [['Steady-state target annual run-rate GBP',null],['One-time cost to separate (one-off + TSA)',null]];
dash.getRange('E9:E10').formulas = [['=Cost_Synergy!I22'],['=Cost_Synergy!I31']];
dash.getRange('E9:E10').format.numberFormat = '#,##0';
dash.getRange('D9:E10').format = dash.getRange('D8:E8').format;
dash.getRange('E9:E10').format.numberFormat = '#,##0';
dash.getRange('L12:M32').unmerge();
dash.getRange('L12:M32').clear({applyTo:'contents'});
dash.getRange('L12:M17').values = [['Bridge A | Annual run-rate (synthetic)','GBP / year'],['Allocated baseline',null],['+ Dis-synergy',null],['Standalone run-rate',null],['- Identified annual synergies',null],['Target annual run-rate',null]];
dash.getRange('M13:M17').formulas = [['=Cost_Synergy!I18'],['=Cost_Synergy!I19'],['=Cost_Synergy!I20'],['=Cost_Synergy!I21'],['=Cost_Synergy!I22']];
dash.getRange('L19:M19').values = [['Bridge B | One-time costs (synthetic)','GBP']];
dash.getRange('L20:L25').values = [['One-off Day 1'],['One-off Wave 2'],['One-off Wave 3'],['Buyer TSA fees'],['One-off separation cost'],['One-time cost to separate (one-off + TSA)']];
dash.getRange('M20:M25').formulas = [['=Cost_Synergy!I26'],['=Cost_Synergy!I27'],['=Cost_Synergy!I28'],['=Cost_Synergy!I29'],['=Cost_Synergy!I30'],['=Cost_Synergy!I31']];
dash.getRange('L27:M27').merge();
dash.getRange('L27').values = [['TSA fees replace standalone run cost for covered services during the TSA period; they are not added to the run-rate.']];
dash.getRange('L12:M12').format = {fill:'#1F3864',font:{name:'Calibri',size:9,bold:true,color:'#FFFFFF'},wrapText:true};
dash.getRange('L19:M19').format = {fill:'#1F3864',font:{name:'Calibri',size:9,bold:true,color:'#FFFFFF'},wrapText:true};
dash.getRange('L13:M17').format.numberFormat = '#,##0';
dash.getRange('L20:M25').format.numberFormat = '#,##0';
dash.getRange('L17:M17').format.font = {name:'Calibri',size:10,bold:true,color:'#000000'};
dash.getRange('L24:M25').format.font = {name:'Calibri',size:10,bold:true,color:'#000000'};
dash.getRange('L27').format = {font:{name:'Calibri',size:8,italic:true,color:'#667085'},wrapText:true};
dash.getRange('L:L').format.columnWidth = 39;
dash.getRange('M:M').format.columnWidth = 18;
dash.getRange('L12:M12').format.rowHeight = 30;
dash.getRange('L19:M19').format.rowHeight = 30;
dash.getRange('L27:M27').format.rowHeight = 36;
dash.deleteAllDrawings();
dash.charts.add('bar',dash.getRange('A12:B18')).setPosition('D11','J19');
dash.charts.add('line',dash.getRange('A21:B45')).setPosition('D21','J40');
dash.charts.add('bar',dash.getRange('L12:M17')).setPosition('D47','J63');
dash.charts.add('bar',dash.getRange('L19:M25')).setPosition('D65','J81');

// Update assumption narrative and make precedence and override explicit.
const assumptions = wb.worksheets.getItem('README_Assumptions');
assumptions.getRange('B8').values = [['Disposition precedence: shared with seller = TSA-then-replace; else transferable Y and Low entanglement = Transfer; else Consent = Clone; else High entanglement = Replicate; otherwise Replace. Final disposition uses a nonblank override before the suggested disposition. APP010 has the single illustrative Retire override; Retire is not inferred from source data.']];
assumptions.getRange('B8').format.wrapText = true;
assumptions.getRange('A8:E8').format.rowHeight = 70;
assumptions.getRange('A39').values = [['Override example: APP010, criticality 2, suggested Clone, final Retire. Synthetic rationale: function absorbed by buyer platform. Every override requires a rationale.']];
assumptions.getRange('A39:E39').merge();
assumptions.getRange('A39:E39').format = {font:{name:'Calibri',size:9,italic:true,color:'#667085'},wrapText:true};
assumptions.getRange('B40').values = [['Bridge A is an annual steady-state run-rate bridge: allocated baseline + dis-synergy = standalone run-rate, then subtract identified annual midpoint synergies to reach target run-rate. Bridge B totals one-off Day 1, Wave 2 and Wave 3 costs plus buyer-side planned TSA fees through exit. TSA fees replace standalone run cost for covered services during the TSA period; they are not added to the run-rate.']];
assumptions.getRange('A40:E40').merge();
assumptions.getRange('A40:E40').format = {font:{name:'Calibri',size:9,italic:true,color:'#667085'},wrapText:true};
assumptions.getRange('A41:E41').merge();
assumptions.getRange('A41').values = [['Synthetic actuals are tracker-only; planned and identified synergy is not represented as achieved.']];
assumptions.getRange('A41:E41').format = {font:{name:'Calibri',size:9,italic:true,color:'#667085'},wrapText:true};
assumptions.getRange('A42:E42').merge();
assumptions.getRange('A42').values = [['An override changes disposition and counts; the one-off cost estimate is an owner input and must be re-estimated separately.']];
assumptions.getRange('A42:E42').format = {font:{name:'Calibri',size:9,italic:true,color:'#667085'},wrapText:true};

wb.recalculate();
cost.getRange('I26').formulas = [['=G7']];
cost.getRange('I27').formulas = [['=G6']];
wb.recalculate();
for (const [sheetName,range,name] of [['App_Disposition','A1:T16','app-after'],['Cost_Synergy','A15:U36','cost-after'],['Dashboard','A1:M85','dash-after'],['Checks','A1:D15','checks-after']]) {
  const preview = await wb.render({sheetName,range,scale:1,format:'png'});
  await fs.writeFile(`${root}/.build_deck/${name}.png`,new Uint8Array(await preview.arrayBuffer()));
}
console.log((await wb.inspect({kind:'region',sheetId:'Checks',range:'A4:C13',maxChars:3500})).ndjson);
console.log('override/final/suggested/wave/tsa',JSON.stringify(app.getRange('A14:T14').values),JSON.stringify(app.getRange('O14:T14').formulas));
console.log('bridge values',JSON.stringify(cost.getRange('H16:I33').values),JSON.stringify(cost.getRange('I18:I31').formulas));
await (await SpreadsheetFile.exportXlsx(wb)).save(xlsx);
