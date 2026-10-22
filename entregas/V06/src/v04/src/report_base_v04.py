from engine import *
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import html,json,hashlib,openpyxl,re
OUT=ROOT/'entrega'
def n(x,d=2):
 if x is None:return 'Não aplicável'
 return f'{x:,.{d}f}'.replace(',','X').replace('.',',').replace('X','.')
def e(x):return n(x)+' €'
def pc(x):return n(x*100)+'%' if x is not None else 'Não aplicável'
def human(content):
 content=str(content)
 content=re.sub(r'([a-zà-ÿ]{2,})(\d)',r'\1 \2',content)
 content=re.sub(r'(\d)([A-Za-zÀ-ÿ]{2,})',r'\1 \2',content)
 content=re.sub(r'([,;])(?=[A-Za-zÀ-ÿ0-9])',r'\1 ',content)
 for word in ['CAPEX','EBITDA','CFADS','GOP','VGV','DSCR','TIR','VPL','EAC','BAC','ETC','CPCV','KK','HMA','LP','JSCV','JALS','Board','Hotel']:
  content=re.sub(r'(?<=[a-zà-ÿ])'+word,' '+word,content)
  content=re.sub(re.escape(word)+r'(?=[a-zà-ÿ0-9])',word+' ',content)
 content=re.sub(r'(?<=[a-zà-ÿ])(?=V(?:0[1-4]|45))',' ',content)
 content=re.sub(r'(?<=[a-zà-ÿ])(?=M\d)',' ',content)
 return content
def report(name,title,blocks):
 styles=getSampleStyleSheet();styles['BodyText'].fontSize=9;styles['BodyText'].leading=13;styles['BodyText'].spaceAfter=7
 story=[Paragraph(html.escape(title),styles['Title'])];md=['# '+title,'','18/10/2026 | KK Hotelaria e Turismo, S.A. | V04','']
 for typ,content in blocks:
  if typ in ['p','h']:
   content=human(content)
   content=re.sub(r'([a-zà-ÿ]{2,})(\d)',r'\1 \2',content)
   content=re.sub(r'(\d)([A-Za-zÀ-ÿ]{2,})',r'\1 \2',content)
   content=re.sub(r'(\d)(M€)',r'\1 \2',content)
   content=re.sub(r'(%)([A-Za-zÀ-ÿ])',r'\1 \2',content)
  if typ=='h':story+=[Spacer(1,8),Paragraph(html.escape(content),styles['Heading2'])];md+=['## '+content,'']
  elif typ=='p':story.append(Paragraph(html.escape(content),styles['BodyText']));md+=[content,'']
  elif typ=='t':
   data=[[Paragraph(html.escape(human(c)),styles['BodyText']) for c in row] for row in content];cols=len(data[0]);widths=[490/cols]*cols
   if cols==4:widths=[170,106,106,108]
   t=Table(data,colWidths=widths,repeatRows=1,hAlign='LEFT');t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#DFEAF1')),('VALIGN',(0,0),(-1,-1),'TOP'),('GRID',(0,0),(-1,-1),.2,colors.lightgrey),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5)]));story+=[t,Spacer(1,10)]
   md+=['| '+' | '.join(map(str,content[0]))+' |','|'+'|'.join(['---']*cols)+'|']+['| '+' | '.join(map(str,row))+' |' for row in content[1:]]+['']
 def foot(c,doc):c.setFont('Helvetica',8);c.setFillColor(colors.HexColor('#526575'));c.drawString(42,25,'KK-01 V04 | Previsões e hipóteses — condições externas pendentes');c.drawRightString(553,25,str(doc.page))
 SimpleDocTemplate(str(OUT/(name+'.pdf')),pagesize=(595,842),leftMargin=42,rightMargin=42,topMargin=42,bottomMargin=42).build(story,onFirstPage=foot,onLaterPages=foot)
 (OUT/(name+'.md')).write_text('\n'.join(md))

