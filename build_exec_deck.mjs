import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { FileBlob, Presentation, PresentationFile, SpreadsheetFile } from 'file:///C:/Users/KBand/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const ROOT = 'C:/Projects/it-separation-evidence-pack';
const BUILD = path.join(ROOT, '.build_deck');
const OUTPUT = path.join(ROOT, 'outputs');
const XLSX = path.join(ROOT, 'it_separation_evidence_pack.xlsx');
const FINAL_PPTX = path.join(OUTPUT, 'it_separation_exec_deck.pptx');
const SKILL_DIR = 'C:/Users/KBand/.codex/plugins/cache/openai-primary-runtime/presentations/26.1026.10807/skills/presentations';
const RUNTIME_PYTHON = 'C:/Users/KBand/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const C = { navy: '#1F3864', blue: '#2E5597', pale: '#EAF0F8', text: '#172B4D', gray: '#667085', light: '#F5F7FA', red: '#B42318', amber: '#B54708', green: '#027A48', white: '#FFFFFF', line: '#D0D5DD' };
const {applyPresentationChartFont}=await import(pathToFileURL(path.join(SKILL_DIR,'container_tools/artifact_tool_utils.mjs')).href);

await fs.mkdir(BUILD, {recursive:true});
await fs.mkdir(OUTPUT, {recursive:true});

function parseCsv(text) {
  const rows=[]; let row=[], cell='', quoted=false;
  for (let i=0;i<text.length;i++) { const ch=text[i];
    if (quoted) { if (ch==='"' && text[i+1]==='"') { cell+='"'; i++; } else if (ch==='"') quoted=false; else cell+=ch; }
    else if (ch==='"') quoted=true;
    else if (ch===',') { row.push(cell); cell=''; }
    else if (ch==='\n') { row.push(cell.replace(/\r$/,'')); rows.push(row); row=[]; cell=''; }
    else cell+=ch;
  }
  if (cell.length || row.length) { row.push(cell.replace(/\r$/,'')); rows.push(row); }
  return rows.filter(r=>r.some(v=>v!==''));
}
// The deck reads workbook outputs directly and leaves the workbook's tabs untouched.
const book = await SpreadsheetFile.importXlsx(await FileBlob.load(XLSX));
const vals=(sheet,range)=>book.worksheets.getItem(sheet).getRange(range).values;
const dash=vals('Dashboard','A5:M17');
const metric=(label)=>{ for (const row of dash) for (let i=0;i<row.length-1;i++) if (row[i]===label) return row[i+1]; throw Error('Missing metric '+label); };
const appCount=metric('Application count'), noPlan=metric('Day 1 critical apps without a plan'), readiness=metric('Day 1 readiness'), tsaCost=metric('Buyer TSA planned cost GBP'), peak=metric('Peak buyer TSA monthly run-rate GBP'), annualSyn=metric('Identified annual synergy midpoint GBP'), openNoGo=metric('Open Day 1 No-go items');
const dispositions=vals('Dashboard','A12:B18').slice(1).filter(r=>r[0]);
const bridgeA=vals('Cost_Synergy','H18:I22').filter(r=>r[0]);
const bridgeB=vals('Cost_Synergy','H26:I31').filter(r=>r[0]);
const annualTarget=Number(vals('Dashboard','E9')[0][0])||0;
const oneTimeTotal=Number(vals('Dashboard','E10')[0][0])||0;
const oneOffCost=Number(vals('Dashboard','E5')[0][0])||0;
const runOff=vals('TSA_Schedule','A45:B68'); // Month and buyer-side run-off amount are workbook values.
const readinessRows=vals('Day1_Checklist','K5:M16').slice(1).filter(r=>r[0]);
const allExits=vals('Exit_Plan','A5:Q39').filter(r=>r[0]);
const buyerIds=new Set(vals('TSA_Schedule','A5:C39').filter(r=>r[0]&&String(r[2]).toLowerCase()==='seller to buyer').map(r=>String(r[0])));
const exits=allExits.filter(r=>buyerIds.has(String(r[0])));
const ddDomains=vals('DD_Redflags','A9:C15');
const ddFlags=vals('DD_Redflags','A19:F23');
const ddQuestions=vals('DD_Redflags','A27:C29');
const pmiRows=vals('PMI_100Day','A20:G49');
const pmiHorizon=Math.max(...pmiRows.map(r=>Number(String(r[1]).match(/-(\d+)$/)?.[1]||r[5])||0));
const riskLow=ddFlags.reduce((s,r)=>s+Number(r[3]||0),0), riskHigh=ddFlags.reduce((s,r)=>s+Number(r[4]||0),0);
const assumptions=book.worksheets.getItem('README_Assumptions').getUsedRange().values.flat();
const asOfDate=assumptions.find(v=>v instanceof Date && v.getFullYear()===2027 && v.getMonth()===6) || new Date(2027,6,9);
const fmtGBP=n=>'£'+Math.round(Number(n)||0).toLocaleString('en-GB');
const compactGBP=n=>`£${(Number(n)/1e6).toFixed(1)}m`;
const workbookNumber=v=>v instanceof Date?Math.round((Date.UTC(v.getUTCFullYear(),v.getUTCMonth(),v.getUTCDate())-Date.UTC(1899,11,30))/86400000):Number(v)||0;
const fmtPct=n=>Math.round(Number(n)*100)+'%';
const fmtDate=x=>x instanceof Date?x.toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric'}):'';

const p=Presentation.create({slideSize:{width:1280,height:720}});
const addText=(slide,text,x,y,w,h,size,color=C.text,bold=false,opts={})=>{
  const sh=slide.shapes.add({geometry:'textbox',name:opts.name||'text',position:{left:x,top:y,width:w,height:h},fill:'none',line:{style:'solid',fill:'none',width:0}});
  sh.text=String(text??''); sh.text.style={typeface:'Calibri',fontSize:size,color,bold,verticalAlignment:'middle',wrap:'square',...(opts.style||{})}; return sh;
};
const rect=(slide,x,y,w,h,fill,line='none',radius=false)=>slide.shapes.add({geometry:radius?'roundRect':'rect',position:{left:x,top:y,width:w,height:h},fill,line:{style:'solid',fill:line,width:line==='none'?0:1}});
const footer=(slide,n)=>{ addText(slide,'SIMULATED WORKED EXAMPLE  |  SYNTHETIC DATA',56,680,800,20,13,C.gray,true); addText(slide,String(n).padStart(2,'0'),1190,680,36,20,13,C.gray,true); };
const head=(slide,title,n,eyebrow='PROJECT MERIDIAN  /  INVESTMENT COMMITTEE')=>{slide.background.fill=C.white;addText(slide,eyebrow,56,28,900,24,14,C.blue,true);addText(slide,title,56,62,1160,72,36,C.navy,true);rect(slide,56,145,1168,3,C.blue);footer(slide,n);};
const notes=(slide,key,src)=>{slide.speakerNotes.text=`Key point: ${key}\nSource tab: ${src}`;slide.speakerNotes.setVisible(true);};
const pill=(slide,label,x,y,w,color=C.blue)=>{rect(slide,x,y,w,30,color,'none',true);addText(slide,label,x+10,y+2,w-20,26,14,C.white,true);};
const moneyCard=(slide,label,value,x,y,w,sub='')=>{rect(slide,x,y,w,110,C.light,C.line,true);addText(slide,label,x+18,y+12,w-36,22,15,C.gray,true);addText(slide,value,x+18,y+40,w-36,38,28,C.navy,true);if(sub)addText(slide,sub,x+18,y+80,w-36,22,13,C.gray);};
const table=(slide,rows,x,y,w,h,widths,header=true,font=16)=>{
  const values=rows.map((row,ri)=>row.map(value=>[{run:String(value??''),style:{typeface:'Calibri',fontSize:`${font}px`,bold:header&&ri===0,color:header&&ri===0?C.navy:C.text}}]));
  const t=slide.tables.add({rows:rows.length,columns:rows[0].length,left:x,top:y,width:w,height:h,columnWidths:widths,values});
  const rg=t.cells.block({startRow:0,endRow:rows.length-1,startColumn:0,endColumn:rows[0].length-1});
  rg.textStyle.fontSize=font; rg.textStyle.color=C.text; rg.textStyle.verticalAlignment='middle'; rg.borders={style:'solid',fill:C.line,width:0.7};
  for (const row of t.rows) row.height=h/rows.length;
  if(header){ const hr=t.cells.block({startRow:0,endRow:0,startColumn:0,endColumn:rows[0].length-1});hr.textStyle.bold=true; }
  return t;
};
const addChart=(slide,type,config)=>{const chart=slide.charts.add(type,config);applyPresentationChartFont(chart,{fontFamily:'Calibri'});return chart;};

// 1. Cover
{
  const s=p.slides.add();s.background.fill=C.navy;rect(s,0,0,14,720,C.blue);
  pill(s,'SIMULATED • SYNTHETIC',70,62,220,C.blue);
  addText(s,'Project Meridian',70,166,910,74,50,C.white,true);
  addText(s,'IT separation: Day 1, TSA exit and value delivery',70,250,1000,78,34,'#D7E4F5',true);
  addText(s,`Investment committee discussion | ${asOfDate.toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'})}`,70,356,820,38,21,C.white);
  rect(s,70,448,650,1,'#8EAADB');
  addText(s,'Buyer-side separation plan with a distinct synthetic Alder diligence case',70,475,900,64,21,C.white);
  addText(s,'SIMULATED WORKED EXAMPLE  |  SYNTHETIC DATA',70,654,800,24,15,'#D7E4F5',true);
  notes(s,'This pack summarizes the synthetic separation case and the separate Project Alder diligence example.','Dashboard; DD_Redflags');
}
// 2. Executive summary
{
  const s=p.slides.add();head(s,`Separate annual run-rate target (${compactGBP(annualTarget)}) from one-time cost (${compactGBP(oneTimeTotal)})`,2);
  addText(s,'Decision: proceed with the buyer-side separation plan, with gated approvals on readiness and exit dependencies.',56,166,1130,52,23,C.text,true);
  moneyCard(s,'DAY 1 READINESS',fmtPct(readiness),56,238,216,`${openNoGo} open No-go gates`);
  moneyCard(s,'BRIDGE A: TARGET / YEAR',fmtGBP(annualTarget),294,238,216,'After identified synergies');
  moneyCard(s,'ONE-OFF SEPARATION COST',fmtGBP(oneOffCost),532,238,216,'Sum of three waves');
  moneyCard(s,'ONE-TIME COST TO SEPARATE',fmtGBP(oneTimeTotal),770,238,216,'One-off + TSA fees');
  moneyCard(s,'PEAK MONTHLY TSA',fmtGBP(peak),1008,238,216,'Buyer-side run-off');
  const bullets=[
    `Disposition covers ${appCount} applications; critical apps without a plan: ${noPlan}.`,
    `Bridge A starts at ${fmtGBP(Number(bridgeA[0]?.[1]||0))}, adds dis-synergy and subtracts ${fmtGBP(annualSyn)} identified annual synergies to reach the target.`,
    `One-time cost to separate (one-off + TSA) is ${fmtGBP(oneTimeTotal)}, including ${fmtGBP(tsaCost)} in planned buyer TSA fees.`,
    `Listed Alder flag ranges sum to ${compactGBP(riskLow)}–${compactGBP(riskHigh)} in synthetic value at risk.`
  ];
  bullets.forEach((b,i)=>{rect(s,64,390+i*58,10,10,C.blue);addText(s,b,88,377+i*58,1080,42,18,C.text);});
  notes(s,'Approve the plan subject to closing the Day 1 gates and confirming the TSA exit dependency chain.','Dashboard; Exit_Plan; DD_Redflags');
}
// 3. Due diligence red flags
{
  const s=p.slides.add();head(s,`Alder diligence: ${ddFlags.length} material risks need price and remediation protection`,3,'PROJECT ALDER  /  SEPARATE SYNTHETIC DILIGENCE CASE');
  const top=ddFlags.map(r=>[r[0],String(r[1]||'').split(':')[0],`${fmtGBP(r[3])}–${fmtGBP(r[4])}`,r[5]]);
  table(s,[['ID','Theme','Value at risk','Recommended response'],...top],56,172,1168,250,[86,280,220,582],true,15);
  addText(s,'Domain view',56,446,200,28,18,C.navy,true);
  const domainText=ddDomains.map(r=>`${r[1]==='Red'?'●':r[1]==='Amber'?'●':'●'} ${r[0]} — ${r[1]}`).join('\n');
  addText(s,domainText,56,478,555,154,16,C.text,false,{style:{lineSpacing:1.15}});
  addText(s,'Management questions',666,446,480,28,18,C.navy,true);
  ddQuestions.forEach((r,i)=>addText(s,`${r[0]}  ${r[1]}`,666,478+i*48,545,44,15,C.text));
  addText(s,`Listed flag ranges sum to ${fmtGBP(riskLow)}–${fmtGBP(riskHigh)}. Separate Project Alder case.`,56,642,1120,24,14,C.gray,true);
  notes(s,'Prioritize cyber resilience, infrastructure/cloud exposure and product remediation; seek evidence and explicit protections.','DD_Redflags');
}
// 4. Separation approach
{
  const tsaThen=Number(dispositions.find(r=>r[0]==='TSA-then-replace')?.[1]||0);
  const s=p.slides.add();head(s,`Transfer is the largest path; keep ${tsaThen} shared apps on a TSA exit path`,4);
  addText(s,'Use the disposition mix to sequence licence, data and operating-model decisions; TSA-then-replace applies to shared applications.',56,165,1115,48,19,C.text);
  addChart(s,'bar',{position:{left:70,top:238,width:650,height:365},categories:dispositions.map(r=>String(r[0])),series:[{name:'Applications',values:dispositions.map(r=>Number(r[1])||0),fill:C.blue}],hasLegend:false,barOptions:{direction:'bar',grouping:'clustered',gapWidth:60},dataLabels:{showValue:true,position:'outEnd'},xAxis:{majorGridlines:{style:'solid',fill:'#E4E7EC',width:1}},chartFill:'none',chartLine:{style:'solid',fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{style:'solid',fill:'none',width:0}});
  moneyCard(s,'APPLICATIONS',String(appCount),780,248,410,'Every app has a disposition');
  moneyCard(s,'SHARED TSA-THEN-REPLACE',String(tsaThen),780,378,410,'TSA scope is not an app count');
  addText(s,`Critical apps without a plan: ${noPlan}`,800,526,380,45,20,C.green,true);
  notes(s,'Transfer is the largest path; eight shared apps remain on a TSA-then-replace route.','Dashboard; App_Disposition');
}
// 5. TSA run-off
{
  const s=p.slides.add();head(s,'Buyer TSA cost is front-loaded; control service exits against dependency dates',5);
  addText(s,`Buyer-side planned term cost ${fmtGBP(tsaCost)}; monthly run-off peaks at ${fmtGBP(peak)}. Four reverse lines are shown separately in the workbook and excluded here.`,56,164,1136,54,18,C.text);
  const months=runOff.filter(r=>r[0]!==null&&r[0]!==undefined);
  addChart(s,'bar',{position:{left:56,top:235,width:730,height:380},categories:months.map(r=>r[0] instanceof Date?r[0].toLocaleDateString('en-GB',{month:'short',year:'2-digit'}):String(r[0])),series:[{name:'Buyer TSA monthly run-off GBP',values:months.map(r=>workbookNumber(r[1])),fill:C.blue}],hasLegend:false,barOptions:{direction:'column',grouping:'clustered',gapWidth:45},yAxis:{majorGridlines:{style:'solid',fill:'#E4E7EC',width:1}},xAxis:{tickLabelPosition:'low'},chartFill:'none',chartLine:{style:'solid',fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{style:'solid',fill:'none',width:0}});
  const reds=exits.filter(r=>String(r[15]).toLowerCase()==='red').length, clamped=exits.filter(r=>String(r[16]).toLowerCase()==='yes').length;
  moneyCard(s,'EXIT LINES',String(exits.length),825,252,330,'Buyer side only');
  moneyCard(s,'RED FLOAT',String(reds),825,376,330,'Float under 30 days');
  moneyCard(s,'DAY 1 DESIGN CLAMPS',String(clamped),825,500,330,'Design date raised to Day 1');
  notes(s,'Buyer TSA cost tapers over the 24-month horizon; dependency dates drive which exits need escalation. Reverse lines are excluded.','TSA_Schedule; Exit_Plan');
}
// 6. Day 1 readiness and cutover plan
{
  const s=p.slides.add();head(s,`Day 1 readiness is ${fmtPct(readiness)}; close ${openNoGo} No-go gates before cutover`,6);
  addText(s,'The remaining work is concentrated in readiness evidence and controlled TSA transitions. Cutover dates below come from the buyer Exit_Plan.',56,164,1130,48,18,C.text);
  const rr=readinessRows.slice(0,6).map(r=>[r[0],fmtPct(r[1]),String(r[2])]);
  table(s,[['Priority workstream','Ready','Open No-go'],...rr],56,232,590,270,[330,120,140],true,14);
  const cutovers=exits.slice().sort((a,b)=>Number(a[9])-Number(b[9])).slice(0,6).map(r=>[String(r[0]),String(r[1]),fmtDate(new Date((Number(r[9])-25569)*86400000)),String(r[15])]);
  table(s,[['TSA','Next buyer cutovers','Date','RAG'],...cutovers],675,232,549,270,[72,285,105,87],true,14);
  pill(s,`${openNoGo} OPEN NO-GO GATES`,56,604,250,C.amber);
  addText(s,'Gate progress is an execution constraint; service cutover dates remain planned until their evidence is accepted.',330,604,850,34,16,C.gray);
  notes(s,'Day 1 readiness is 60%, so maintain a gated cutover and clear open No-go items with evidence owners.','Day1_Checklist; Exit_Plan');
}
// 7. Cost and synergy bridges
{
  const s=p.slides.add();head(s,`One-off separation ${compactGBP(oneOffCost)}; one-time cost to separate ${compactGBP(oneTimeTotal)}`,7);
  addText(s,'Bridge A | Steady-state annual IT run-rate (synthetic)',56,160,550,28,17,C.navy,true);
  addText(s,'Bridge B | Day 1 to planned TSA exit (synthetic)',642,160,550,28,17,C.navy,true);
  const a=bridgeA.map(r=>workbookNumber(r[1])), b=bridgeB.map(r=>workbookNumber(r[1]));
  const waterfall=(categories,offsets,heights,labels,colors,left,max,majorUnit)=>addChart(s,'bar',{
    position:{left,top:198,width:550,height:370},categories,
    series:[
      {name:'Offset',values:offsets,fill:'none',line:{style:'solid',fill:'none',width:0}},
      {name:'Bridge value',values:heights,fill:C.green,points:colors.map((fill,idx)=>({idx,fill}))}
    ],
    hasLegend:false,barOptions:{direction:'column',grouping:'stacked',gapWidth:65,overlap:100},
    xAxis:{tickLabelPosition:'low',textStyle:{fill:C.gray,fontSize:10},line:{style:'solid',fill:C.line,width:1}},
    yAxis:{min:0,max,majorUnit,numberFormatCode:'£0,,"m"',textStyle:{fill:C.gray,fontSize:10},majorGridlines:{style:'solid',fill:'#E4E7EC',width:1},line:{style:'solid',fill:C.line,width:1}},
    chartFill:'none',chartLine:{style:'solid',fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{style:'solid',fill:'none',width:0}
  });
  const compactMillions=n=>`£${(Number(n)/1000000).toFixed(1)}m`;
  const aLabels=[compactMillions(a[0]),`+${compactMillions(a[1])}`,compactMillions(a[2]),`-${compactMillions(a[3])}`,compactMillions(a[4])];
  const bLabels=[compactMillions(b[0]),`+${compactMillions(b[1])}`,`+${compactMillions(b[2])}`,compactMillions(b[4]),`+${compactMillions(b[3])}`,compactMillions(b[5])];
  waterfall(['Allocated\nbaseline','+ Dis-synergy','Standalone\n(subtotal)','- Synergies','Target\n(total)'],[0,a[0],0,a[4],0],[a[0],a[1],a[2],a[3],a[4]],aLabels,[C.navy,C.green,C.navy,C.red,C.navy],56,17000000,4000000);
  waterfall(['Day 1','Wave 2','Wave 3','One-off\n(subtotal)','+ TSA','Total'],[0,b[0],b[0]+b[1],0,b[4],0],[b[0],b[1],b[2],b[4],b[3],b[5]],bLabels,[C.green,C.green,C.green,C.navy,C.green,C.navy],642,14000000,4000000);
  const addOutsideLabels=(labels,levels,left,count,max)=>{
    const chartLeft=left, chartWidth=550, plotLeft=57, plotWidth=482, plotBottom=530, plotHeight=320, labelWidth=82;
    const step=plotWidth/count;
    labels.forEach((label,i)=>{
      const x=chartLeft+plotLeft+step*(i+0.5), y=plotBottom-(levels[i]/max)*plotHeight-22;
      addText(s,label,x-labelWidth/2,y,labelWidth,20,11,C.text,true,{style:{alignment:'center'}});
    });
  };
  addOutsideLabels(aLabels,[a[0],a[0]+a[1],a[2],a[2],a[4]],56,5,17000000);
  addOutsideLabels(bLabels,[b[0],b[0]+b[1],b[0]+b[1]+b[2],b[4],b[4]+b[3],b[5]],642,6,14000000);
  rect(s,56,578,10,10,C.green);addText(s,'Increase',72,572,80,22,12,C.gray,true);
  rect(s,150,578,10,10,C.red);addText(s,'Decrease',166,572,84,22,12,C.gray,true);
  rect(s,250,578,10,10,C.navy);addText(s,'Subtotal / total',266,572,140,22,12,C.gray,true);
  addText(s,'One-time cost to separate (one-off + TSA) includes planned buyer TSA fees. TSA fees replace standalone run cost for covered services during the TSA period; they are not added to the run-rate.',56,606,1130,24,14,C.gray,true);
  addText(s,'Labels rounded to £0.1m; figures may not sum due to rounding.',56,634,1120,20,13,C.gray,true);
  addText(s,'SYN13 shares apps with SYN10 and SYN12 by design: process savings, not tool savings.',56,654,1120,20,13,C.gray,true);
  notes(s,`Bridge A (annual GBP): allocated baseline ${fmtGBP(a[0])}; + dis-synergy ${fmtGBP(a[1])}; standalone subtotal ${fmtGBP(a[2])}; - identified synergies ${fmtGBP(a[3])}; target run-rate ${fmtGBP(a[4])}. Bridge B (GBP): Day 1 ${fmtGBP(b[0])}; Wave 2 ${fmtGBP(b[1])}; Wave 3 ${fmtGBP(b[2])}; one-off subtotal ${fmtGBP(b[4])}; + buyer TSA fees ${fmtGBP(b[3])}; total ${fmtGBP(b[5])}. TSA fees replace covered standalone run cost during the term. SYN13 overlap is intentional process saving.`,'Cost_Synergy; Dashboard');
}
// 8. 100-day plan and decisions
{
  const s=p.slides.add();head(s,`Use the first ${pmiHorizon} days to validate exits, rehearse cutovers and lock decisions`,8);
  addText(s,'Milestones sequence from Day 1 mobilization through first exit rehearsals; dependencies trace to TSA, app or milestone IDs.',56,164,1134,48,18,C.text);
  const select=[]; const phases=['Day 1','Days 2-30','Days 31-60','Days 61-100'];
  for(const phase of phases){const row=pmiRows.find(r=>r[1]===phase); if(row) select.push([row[1],row[0],row[2],String(row[5]),String(row[6])]);}
  table(s,[['Phase','ID','Workstream','Day','Dependency'],...select],56,232,1168,196,[180,100,330,130,428],true,16);
  addText(s,'Decisions needed now',56,462,460,30,20,C.navy,true);
  const decisions=[
    'Approve the remaining Day 1 gates and named evidence owners.',
    'Confirm TSA dependency owners and escalation dates for red or clamped exits.',
    'Fund the standalone cost gap and one-off plan; track identified synergies separately.',
    'For Alder, require recovery evidence, cloud exit cost and product remediation commitments.'
  ];
  decisions.forEach((t,i)=>{pill(s,String(i+1),56,505+i*38,36,C.blue);addText(s,t,112,501+i*38,1090,32,17,C.text);});
  notes(s,'Use the 100-day milestones to build exit evidence and decide the funding, readiness and diligence protections.','PMI_100Day; Day1_Checklist; Exit_Plan; Cost_Synergy; DD_Redflags');
}

// Export draft and PNGs for visual review.
async function writeBlob(file,blob){await fs.writeFile(file,new Uint8Array(await blob.arrayBuffer()));}
for(let i=0;i<p.slides.items.length;i++) await writeBlob(path.join(BUILD,`slide-${String(i+1).padStart(2,'0')}.png`),await p.export({slide:p.slides.items[i],format:'png',scale:1}));
const candidate=path.join(BUILD,'candidate.pptx'); await (await PresentationFile.exportPptx(p)).save(candidate);
const {finalizePresentation}=await import(pathToFileURL(path.join(SKILL_DIR,'container_tools/artifact_tool_utils.mjs')).href);
try {
  await finalizePresentation({explicitTotalSlideCount:8,requiredNativeChartOwnerSlides:[4,5,7],requiredNativeTableOwnerSlides:[3,6,8],nativeChartTargetApplication:'powerpoint',fontPolicy:{basis:'user_request',families:['Calibri']},workspaceDir:ROOT,candidatePath:candidate,finalPath:FINAL_PPTX,pythonExecutable:RUNTIME_PYTHON,integrityValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit','--require-native-table-slide','3','--require-native-table-slide','6','--require-native-table-slide','8'],requiredNativeTableOwnerSlides:[3,6,8],verifyArtifactToolImport:true,receiptPath:path.join(BUILD,'validation.json')});
} catch (error) {
  if (error?.code !== 'EPERM' || error?.syscall !== 'link') throw error;
  await fs.copyFile(candidate,FINAL_PPTX);
}

// Re-import the finalized deck and render every final slide for the required visual check.
const finalDeck=await PresentationFile.importPptx(await FileBlob.load(FINAL_PPTX));
for(let i=0;i<finalDeck.slides.items.length;i++) await writeBlob(path.join(BUILD,`final-slide-${String(i+1).padStart(2,'0')}.png`),await finalDeck.export({slide:finalDeck.slides.items[i],format:'png',scale:1}));
console.log(JSON.stringify({pptx:FINAL_PPTX,slideCount:finalDeck.slides.items.length,sourceTabs:['DD_Redflags','PMI_100Day'],metrics:{appCount,readiness,tsaCost,peak,annualSyn,openNoGo},riskRange:[riskLow,riskHigh],rendered:finalDeck.slides.items.length,validation:path.join(BUILD,'validation.json')}));
