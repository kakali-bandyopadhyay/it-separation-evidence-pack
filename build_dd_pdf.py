from pathlib import Path
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from pypdf import PdfReader

root=Path(__file__).resolve().parent; w=load_workbook(root/'it_separation_evidence_pack.xlsx',data_only=True); s=w['DD_Redflags']; out=root/'it_dd_redflag_summary.pdf'
navy=colors.HexColor('#1F3864'); blue=colors.HexColor('#2E5597'); pale=colors.HexColor('#EAF0F8')
st=getSampleStyleSheet(); st.add(ParagraphStyle(name='Title2',fontSize=16,leading=18,textColor=navy,spaceAfter=3)); st.add(ParagraphStyle(name='Small2',fontSize=7.2,leading=8.5)); st.add(ParagraphStyle(name='Head2',fontSize=9.5,leading=11,textColor=navy,spaceBefore=3,spaceAfter=2))
P=lambda x,n='Small2':Paragraph(str(x),st[n]); story=[P('Project Alder | IT due diligence red flags','Title2'),P('<b>SIMULATED • SYNTHETIC DATA • WORKED EXAMPLE</b> | Separate target-side diligence case')]
story += [P('Scope and method','Head2'),P(s['B4'].value),P(f"Complexity proxy: {s['B6'].value:.1f} / 5. Reuses App_Disposition complexity method; not an Alder inventory score.")]
story += [P('Headline rating by domain','Head2')]
domains=[[P('Domain'),P('RAG'),P('Finding')]]+[[P(s.cell(r,c).value) for c in range(1,4)] for r in range(9,16)]
def table(data,widths):
 t=Table(data,colWidths=widths,repeatRows=1);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),navy),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.3,colors.HexColor('#CBD5E1')),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,pale]),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2)]));return t
story.append(table(domains,[150,42,620]));story.append(P('Top five red flags','Head2'))
flags=[[P('Finding'),P('Why it matters'),P('Value-at-risk GBP'),P('Action')]]
for r in range(19,24): flags.append([P(s.cell(r,2).value),P(s.cell(r,3).value),P(f"{s.cell(r,4).value:,.0f} – {s.cell(r,5).value:,.0f}"),P(s.cell(r,6).value)])
story.append(table(flags,[280,300,115,117]));story.append(P('Questions for management','Head2'))
story.append(P(' &nbsp; '.join(f"<b>{s.cell(r,1).value}</b> {s.cell(r,2).value}" for r in range(27,30))))
story += [P('Limitations','Head2'),P(s['B31'].value),P('Synthetic findings and value ranges are illustrative and require evidence validation. No real deal or achieved outcome is represented.')]
doc=SimpleDocTemplate(str(out),pagesize=landscape(A4),rightMargin=25,leftMargin=25,topMargin=18,bottomMargin=22)
def footer(c,d): c.setFillColor(colors.HexColor('#64748B'));c.setFont('Helvetica',6.5);c.drawString(25,12,'Simulated worked example | synthetic data | Source: supplied CSVs via workbook');c.drawRightString(landscape(A4)[0]-25,12,'1 / 1')
doc.build(story,onFirstPage=footer); pdf=PdfReader(str(out)); assert len(pdf.pages)==1; text=pdf.pages[0].extract_text() or ''; assert 'SYNTHETIC' in text and 'Project Alder' in text; print('PASS | one page |',out)
