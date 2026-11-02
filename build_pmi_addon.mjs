import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from 'file:///C:/Users/KBand/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const ROOT='C:/Projects/it-separation-evidence-pack';
const XLSX=path.join(ROOT,'it_separation_evidence_pack.xlsx');
const C={navy:'#1F3864',blue:'#2E5597',pale:'#EAF0F8',text:'#172B4D',gray:'#667085',white:'#FFFFFF',line:'#D0D5DD',green:'#E2F0D9',red:'#FCE4D6'};

function parseCsv(text){
  const rows=[];let row=[],cell='',quoted=false;
  for(let i=0;i<text.length;i++){const ch=text[i];if(quoted){if(ch==='"'&&text[i+1]==='"'){cell+='"';i++;}else if(ch==='"')quoted=false;else cell+=ch;}else if(ch==='"')quoted=true;else if(ch===','){row.push(cell);cell='';}else if(ch==='\n'){row.push(cell.replace(/\r$/,''));rows.push(row);row=[];cell='';}else cell+=ch;}
  if(cell.length||row.length){row.push(cell.replace(/\r$/,''));rows.push(row);}return rows.filter(r=>r.some(v=>v!==''));
}
const numericCsv=rows=>rows.map((row,i)=>i===0?row:row.map(v=>v!==''&&/^-?\d+(?:\.\d+)?$/.test(v)?Number(v):v));
const csv=async name=>numericCsv(parseCsv(await fs.readFile(path.join(ROOT,'data',name),'utf8')));
const setValues=(ws,anchor,rows)=>{ws.getRange(anchor.split(':')[0]).getResizedRange(rows.length-1,rows[0].length-1).values=rows;};
const setTitle=(ws,range,text)=>{ws.getRange(range).values=[[text]];ws.getRange(range).format={font:{name:'Calibri',size:11,bold:true,color:C.white},fill:C.navy,wrapText:true,verticalAlignment:'center'};};
const setHeader=(ws,range)=>{ws.getRange(range).format={font:{name:'Calibri',size:10,bold:true,color:C.white},fill:C.blue,wrapText:true,verticalAlignment:'center'};};
const setBody=(ws,range)=>{ws.getRange(range).format={font:{name:'Calibri',size:10,color:C.text},wrapText:true,verticalAlignment:'top'};};

const source={
  pmi:await csv('pmi_milestones.csv'),
  risks:await csv('risks.csv'),
  synergies:await csv('synergies.csv'),
  outsourcing:await csv('outsourcing.csv'),
};
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(XLSX));
try{wb.worksheets.getItem('PMI_100Day').delete();}catch{}
const ws=wb.worksheets.add('PMI_100Day');ws.showGridlines=false;

ws.getRange('A1').values=[['Project Meridian | PMI 100-Day Plan']];
ws.getRange('A1:J1').format={font:{name:'Calibri',size:16,bold:true,color:C.white},fill:C.navy,verticalAlignment:'center'};
ws.getRange('A2').values=[['SYNTHETIC | SIMULATED WORKED EXAMPLE | Source rows retained below; amounts are not additive unless explicitly stated.']];
ws.getRange('A2:J2').format={font:{name:'Calibri',size:10,bold:true,color:C.blue},wrapText:true};

setTitle(ws,'A4:B4','PMI principles & governance');
setValues(ws,'A5:B7',[
  ['Governance area','Operating approach'],
  ['Integration management office','Buyer-side IMO coordinates applications, infrastructure, security, data, vendors, and people/service desk against one exit plan.'],
  ['Decision rights','Workstream owner roles in the milestone plan own evidence and dependency actions; cross-workstream conflicts go to the buyer governance forum.'],
  ['Cadence and evidence','Review milestone dependency RAG, risk mitigations, and decisions through the governance meeting; retain evidence against each milestone.'],
]);
setHeader(ws,'A5:B5');setBody(ws,'A6:B8');

setTitle(ws,'A10:B10','100-day scorecard');
setValues(ws,'A11:B15',[
  ['Metric','Value'],
  ['Milestones on-track proxy (non-red dependency RAG)',''],
  ['Identified synergy vs modeled target',''],
  ['Identified annual synergy midpoint GBP',''],
  ['Open Day 1 No-go items',''],
]);
ws.getRange('B12').formulas=[['=IFERROR((COUNTIF($H$20:$H$49,"Green")+COUNTIF($H$20:$H$49,"Amber"))/COUNTA($A$20:$A$49),0)']];
ws.getRange('B13').formulas=[['=IFERROR(SUM($F$53:$F$67)/SUM(Cost_Synergy!$Q$6:$Q$20),0)']];
ws.getRange('B14').formulas=[['=SUM($F$53:$F$67)']];
ws.getRange('B15').formulas=[['=Dashboard!$E$8']];
setHeader(ws,'A11:B11');setBody(ws,'A12:B15');
ws.getRange('B12:B13').format.numberFormat='0%';
ws.getRange('B14').format.numberFormat='£#,##0';
ws.getRange('A16').values=[['Basis: no completion status is supplied; on-track is a non-red dependency-health proxy. Dependency RAG inherits a directly linked TSA service RAG; a resolved app or milestone dependency is Green. Synergy target is the supplied Cost_Synergy annual midpoint list; no separate target is in the CSVs.']];
ws.getRange('A16:J16').format={font:{name:'Calibri',size:9,italic:true,color:C.gray},wrapText:true};

setTitle(ws,'A18:I18','100-day milestone plan | dependency RAG is not a completion status');
setValues(ws,'A19:I19',[['Milestone ID','Phase','Workstream','Milestone','Owner role','Target day','Dependency ID','Dependency RAG','Dependency resolves?']]);setHeader(ws,'A19:I19');
for(let i=0;i<source.pmi.length-1;i++){
  const row=20+i, sourceRow=96+i;
  for(let col=0;col<7;col++)ws.getRange(`${String.fromCharCode(65+col)}${row}`).formulas=[[`=PMI_100Day!${String.fromCharCode(65+col)}${sourceRow}`]];
  ws.getRange(`I${row}`).formulas=[[`=IF(OR(COUNTIF(Exit_Plan!$A$5:$A$39,G${row})>0,COUNTIF(TSA_Schedule!$A$5:$A$39,G${row})>0,COUNTIF(App_Disposition!$A$5:$A$64,G${row})>0,COUNTIF($A$20:$A$49,G${row})>0),"PASS","FAIL")`]];
  ws.getRange(`H${row}`).formulas=[[`=IF(I${row}<>"PASS","Red",IF(COUNTIF(Exit_Plan!$A$5:$A$39,G${row})>0,INDEX(Exit_Plan!$P$5:$P$39,MATCH(G${row},Exit_Plan!$A$5:$A$39,0)),"Green"))`]];
}
setBody(ws,'A20:I49');ws.getRange('F20:F49').format.numberFormat='0';
ws.getRange('H20:I49').format={font:{name:'Calibri',size:10,color:C.text},horizontalAlignment:'center',verticalAlignment:'center'};

setTitle(ws,'A51:J51','Identified synergy and vendor-consolidation levers | linked to Cost_Synergy');
setValues(ws,'A52:J52',[['Synergy ID','Lever','Tower','Linked app IDs','Owner role','Annual midpoint GBP','Confidence %','Start month','Ramp months','Tie check']]);setHeader(ws,'A52:J52');
for(let i=0;i<source.synergies.length-1;i++){
  const row=53+i,cs=6+i;
  const cols=['J','K','L','M','N','Q','T','O','P'];
  cols.forEach((c,j)=>ws.getRange(`${String.fromCharCode(65+j)}${row}`).formulas=[[`=Cost_Synergy!${c}${cs}`]]);
  ws.getRange(`J${row}`).formulas=[[`=IF(AND(A${row}=Cost_Synergy!J${cs},B${row}=Cost_Synergy!K${cs},D${row}=Cost_Synergy!M${cs},F${row}=Cost_Synergy!Q${cs},G${row}=Cost_Synergy!T${cs}),"PASS","FAIL")`]];
}
setBody(ws,'A53:J67');ws.getRange('F53:F67').format.numberFormat='£#,##0';ws.getRange('G53:G67').format.numberFormat='0%';
ws.getRange('A68').values=[['Identified values only; the source files do not supply realized results. SYN13 overlaps app IDs with SYN10 and SYN12 by design (process saving).']];ws.getRange('A68:J68').format={font:{name:'Calibri',size:9,italic:true,color:C.gray},wrapText:true};

setTitle(ws,'A70:F70','Outsourcing options | alternatives, not additive');
setValues(ws,'A71:F71',[['Option ID','Scope','Rationale','Indicative saving low GBP','Indicative saving high GBP','Key risks']]);setHeader(ws,'A71:F71');
for(let i=0;i<source.outsourcing.length-1;i++){
  const row=72+i,sr=163+i;for(let col=0;col<6;col++)ws.getRange(`${String.fromCharCode(65+col)}${row}`).formulas=[[`=PMI_100Day!${String.fromCharCode(65+col)}${sr}`]];
}
setBody(ws,'A72:F76');ws.getRange('D72:E76').format.numberFormat='£#,##0';
ws.getRange('A77').values=[['These options are mutually exclusive alternatives; do not add them to the cost bridge or synergy total.']];ws.getRange('A77:F77').format={font:{name:'Calibri',size:9,italic:true,color:C.gray},wrapText:true};

setTitle(ws,'A79:H79','Top 10 separation risks and mitigations | source-linked');
setValues(ws,'A80:H80',[['Risk ID','Risk','linked_id','Impact','Likelihood','Mitigation','Owner role','Linked ID exists?']]);setHeader(ws,'A80:H80');
for(let i=0;i<source.risks.length-1;i++){
  const row=81+i,sr=130+i;for(let col=0;col<7;col++)ws.getRange(`${String.fromCharCode(65+col)}${row}`).formulas=[[`=PMI_100Day!${String.fromCharCode(65+col)}${sr}`]];
  ws.getRange(`H${row}`).formulas=[[`=IF(OR(COUNTIF(TSA_Schedule!$A$5:$A$39,C${row})>0,COUNTIF(Exit_Plan!$A$5:$A$39,C${row})>0,COUNTIF(App_Disposition!$A$5:$A$64,C${row})>0),"PASS","FAIL")`]];
}
setBody(ws,'A81:H90');ws.getRange('H81:H90').format={font:{name:'Calibri',size:10,color:C.text},horizontalAlignment:'center',verticalAlignment:'center'};

// Source data blocks remain in this tab, unchanged in meaning, for traceability.
const sourceBlocks=[
  [94,'pmi_milestones.csv',source.pmi],
  [128,'risks.csv',source.risks],
  [142,'synergies.csv',source.synergies],
  [161,'outsourcing.csv',source.outsourcing],
];
for(const [row,name,data] of sourceBlocks){ws.getRange(`A${row}`).values=[[`Source rows | data/${name}`]];ws.getRange(`A${row}`).format={font:{name:'Calibri',size:10,bold:true,color:C.navy}};setValues(ws,`A${row+1}`,data);}
// Apply source block header/body styles over each complete CSV table.
for(const [row,,data] of sourceBlocks){
  const end=row+1+data.length-1, width=data[0].length;
  ws.getRange(`A${row+1}`).getResizedRange(0,width-1).format={font:{name:'Calibri',size:9,bold:true,color:C.white},fill:C.blue,wrapText:true};
  ws.getRange(`A${row+2}`).getResizedRange(data.length-2,width-1).format={font:{name:'Calibri',size:9,color:C.gray},wrapText:true,verticalAlignment:'top'};
}

// The Checks tab retains the core pack checks and adds formula-driven add-on A checks.
const checks=wb.worksheets.getItem('Checks');
checks.getRange('A191:C196').values=[
  ['Add-on A | PMI_100Day checks',null,null],
  ['Check','Result','Basis'],
  ['PMI milestone dependency IDs resolve',null,'TSA, app, exit-plan, or source milestone IDs'],
  ['Synergy levers tie to Cost_Synergy',null,'IDs, linked apps, midpoint and confidence'],
  ['Risk linked_id values resolve',null,'risks.csv linked_id against workbook IDs'],
  ['Outsourcing alternatives excluded from bridge/synergy',null,'Options are not additive and are not cost components'],
];
checks.getRange('B193').formulas=[['=IF(COUNTIF(PMI_100Day!$I$20:$I$49,"FAIL")=0,"PASS","FAIL")']];
checks.getRange('B194').formulas=[['=IF(COUNTIF(PMI_100Day!$J$53:$J$67,"FAIL")=0,"PASS","FAIL")']];
checks.getRange('B195').formulas=[['=IF(COUNTIF(PMI_100Day!$H$81:$H$90,"FAIL")=0,"PASS","FAIL")']];
checks.getRange('B196').formulas=[['=IF(AND(COUNTIF(Cost_Synergy!$J$6:$J$20,"OUT*")=0,COUNTIF(Dashboard!$L$13:$L$18,"*outsourcing*")=0),"PASS","FAIL")']];
checks.getRange('A191:C191').format={font:{name:'Calibri',size:11,bold:true,color:C.white},fill:C.navy};
checks.getRange('A192:C192').format={font:{name:'Calibri',size:10,bold:true,color:C.white},fill:C.blue};
checks.getRange('A193:C196').format={font:{name:'Calibri',size:10,color:C.text},wrapText:true};
checks.getRange('B193:B196').format={font:{name:'Calibri',size:10,bold:true,color:C.text}};

// Readable column widths and print-friendly row layouts.
const widths={A:16,B:42,C:26,D:85,E:35,F:12,G:18,H:22,I:22,J:16};
for(const [col,width] of Object.entries(widths))ws.getRange(`${col}:${col}`).format.columnWidth=width;
ws.getRange('A1:J1').format.rowHeight=28;ws.getRange('A2:J2').format.rowHeight=24;
ws.getRange('A16:J16').format.rowHeight=36;ws.getRange('A68:J68').format.rowHeight=30;ws.getRange('A77:F77').format.rowHeight=30;
await (await SpreadsheetFile.exportXlsx(wb)).save(XLSX);
console.log(JSON.stringify({workbook:XLSX,milestones:source.pmi.length-1,risks:source.risks.length-1,synergies:source.synergies.length-1,outsourcingAlternatives:source.outsourcing.length-1,checks:'A193:A196'}));
