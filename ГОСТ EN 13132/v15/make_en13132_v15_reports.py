#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""One-page Word calibration protocol and chart, grounded in the cached audited Excel model.
No synthetic points are added and no absent certificate is silently replaced by zero.
"""
from pathlib import Path
import zipfile, lxml.etree as ET, math
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy.stats import t as student_t
from docx import Document
from docx.shared import Inches,Pt,RGBColor,Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def load(path):
 z=zipfile.ZipFile(path)
 rt={i:ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml')) for i in (1,2,3,6,9,10)}
 def cell(i,code):
  c=rt[i].find(f'.//m:c[@r="{code}"]',NS)
  if c is None:return None
  v=c.find('m:v',NS);it=c.find('m:is',NS)
  if it is not None:return ''.join(it.itertext())
  if v is None:return None
  try:return float(v.text)
  except:return v.text
 a=np.array([cell(1,f'H{i}') for i in range(7,13)],dtype=float)
 y=np.array([cell(1,f'M{i}') for i in range(7,13)],dtype=float)
 return dict(c=cell,x=a,y=y,a=cell(3,'B11'),b=cell(3,'B12'),s2=cell(3,'B14'),n=int(cell(3,'B4')),xbar=cell(3,'B6'),sxx=cell(3,'B8'),r2=cell(3,'B21'),t_b=cell(3,'B19'),b_sig=cell(3,'B20'),f_lof=cell(3,'B26'),f_crit=cell(3,'B27'),lof=cell(3,'B29'),xstar=cell(3,'B31'),ureg=cell(3,'B32'),upurity=cell(3,'B33'),uweights=cell(3,'B34'),umodel=cell(3,'B35'),budget=[cell(6,f'B{i}') for i in (19,20,21,23)],shares=[cell(6,f'D{i}') for i in (19,20,21,22,23)],variance=cell(6,'C25'),weights_mode=cell(9,'B20'),xstar_reason=cell(10,'B9'),status=cell(6,'B6'))

def shade(cell, color):
 tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd');shd.set(qn('w:fill'),color);tcPr.append(shd)

def fnum(v,n=5):return 'н/д' if v is None or not isinstance(v,(float,int,np.floating)) else f'{v:.{n}f}'.replace('.',',')

def make_plot(data,basename):
 x=data['x'];y=data['y']; n=len(x)
 xx=np.linspace(float(min(x)),float(max(x)),240)
 mu=data['a']*xx+data['b']; var=data['s2']*(1/n+(xx-data['xbar'])**2/data['sxx'])
 d=student_t.ppf(.975,n-2)*np.sqrt(var)
 fig,ax=plt.subplots(figsize=(7.6,3.0),dpi=190)
 ax.scatter(x,y, s=27,facecolor='#176B92',edgecolor='#176B92',marker='o',label='Эксперимент',zorder=5)
 ax.plot(xx,mu,color='#2965A1',lw=1.8,label='OLS: средняя линия')
 ax.plot(xx,mu-d,color='#A0C5E3',lw=1.2,linestyle='--',label='95% ДИ (точечный)')
 ax.plot(xx,mu+d,color='#A0C5E3',lw=1.2,linestyle='--')
 ax.set_xlabel('Xэфф = mаналита·pаналита / (mИС·pИС), г/г',fontsize=8)
 ax.set_ylabel('Y = Sаналита / SИС',fontsize=8)
 ax.grid(True,linewidth=.4,alpha=.18)
 ax.tick_params(labelsize=8); ax.set_xlim(min(x)-.035*(max(x)-min(x)),max(x)+.035*(max(x)-min(x)))
 ax.legend(loc='upper left',fontsize=7,frameon=False,ncol=3)
 ax.spines['top'].set_visible(False);ax.spines['right'].set_visible(False)
 fig.tight_layout(pad=.65)
 out=f'/mnt/data/{basename}_график.png';fig.savefig(out,dpi=190);plt.close(fig)
 return out

def document(path,name):
 d=load(path); plot=make_plot(d,f'EN13132_v15_{name}')
 doc=Document();sec=doc.sections[0];sec.page_height=Cm(29.7);sec.page_width=Cm(21)
 sec.top_margin=Cm(1.28);sec.bottom_margin=Cm(1.2);sec.left_margin=Cm(1.65);sec.right_margin=Cm(1.65)
 styles=doc.styles;normal=styles['Normal'];normal.font.name='Arial';normal.font.size=Pt(8.5);normal.font.color.rgb=RGBColor(44,57,68)
 normal.paragraph_format.space_after=Pt(3)
 p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(3)
 r=p.add_run(f'ГРАДУИРОВОЧНАЯ ХАРАКТЕРИСТИКА  |  {name}');r.bold=True;r.font.size=Pt(13);r.font.color.rgb=RGBColor(33,72,104)
 p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(4)
 p.add_run('ГОСТ EN 13132-2012  ·  Весовое приготовление в н-гептане  ·  Диагностическая модель OLS').italic=True
 pic=doc.add_paragraph();pic.alignment=WD_ALIGN_PARAGRAPH.CENTER;pic.paragraph_format.space_after=Pt(4)
 pic.add_run().add_picture(plot,width=Inches(6.75))
 p=doc.add_paragraph();p.paragraph_format.space_after=Pt(3)
 p.add_run('Уравнение: ').bold=True;p.add_run(f'Y = {fnum(d["a"],7)}·X + {fnum(d["b"],7)}    ')
 p.add_run('R² = ').bold=True;p.add_run(fnum(d['r2'],6)+'    ')
 p.add_run('b значим: ').bold=True;p.add_run(str(d['b_sig']))
 p=doc.add_paragraph();p.paragraph_format.space_after=Pt(4)
 p.add_run('Lack-of-fit: ').bold=True;p.add_run(f'F = {fnum(d["f_lof"],4)}; Fкр = {fnum(d["f_crit"],4)}; '+str(d['lof']))
 p=doc.add_paragraph('Экспериментальные уровни и относительный коэффициент fᵢ=Xэфф/Y');p.paragraph_format.space_after=Pt(3)
 p.runs[0].bold=True
 tab=doc.add_table(rows=1,cols=5);tab.alignment=WD_TABLE_ALIGNMENT.CENTER;tab.style='Table Grid'
 heads=['№','Xэфф (г/г)','Y средний','Ŷ OLS','fᵢ=X/Y']
 for j,x in enumerate(heads):
  c=tab.rows[0].cells[j];c.text=x;shade(c,'315A7C')
  for rr in c.paragraphs[0].runs:rr.font.bold=True;rr.font.color.rgb=RGBColor(255,255,255);rr.font.size=Pt(8)
 for i,(x,y) in enumerate(zip(d['x'],d['y']),1):
  row=tab.add_row().cells; vals=[str(i),fnum(x,6),fnum(y,6),fnum(d['a']*x+d['b'],6),fnum(x/y,6)]
  for j,v in enumerate(vals):
   row[j].text=v
   if i%2==0:shade(row[j],'F0F4F8')
   for pr in row[j].paragraphs:
    pr.paragraph_format.space_after=Pt(0)
    for rr in pr.runs:rr.font.size=Pt(7.6)
 doc.add_paragraph().paragraph_format.space_after=Pt(1)
 p=doc.add_paragraph();p.paragraph_format.space_after=Pt(3)
 p.add_run(f'Бюджет в X* = {fnum(d["xstar"],6)} г/г').bold=True
 p.add_run('  ·  оценка только выбранной модели, не полный результат анализа пробы')
 tb=doc.add_table(rows=1,cols=3);tb.style='Table Grid'
 for j,h in enumerate(['Составляющая','u(X*) (г/г)','Доля дисперсии']):
  cell=tb.rows[0].cells[j];cell.text=h;shade(cell,'315A7C')
  for run in cell.paragraphs[0].runs:run.font.bold=True;run.font.color.rgb=RGBColor(255,255,255);run.font.size=Pt(8)
 budget=[('Регрессия (включая cov(a,b))',d['budget'][0],d['shares'][0]),('Чистота аналита',d['budget'][1],d['shares'][1]),('Чистота внутреннего стандарта',d['budget'][2],d['shares'][2]),('Ковариация чистот (знаковая)',None,d['shares'][3]),('Взвешивание (сценарий)',d['budget'][3],d['shares'][4])]
 for i,(label,u,share) in enumerate(budget):
  rw=tb.add_row().cells
  for j,v in enumerate((label, fnum(u,7) if u is not None else 'отдельный член u²',fnum(share,3)+' %')):
   rw[j].text=v
   if i%2==1:shade(rw[j],'F0F4F8')
   for pr in rw[j].paragraphs:
    pr.paragraph_format.space_after=Pt(0)
    for run in pr.runs:run.font.size=Pt(7.5)
 p=doc.add_paragraph();p.paragraph_format.space_before=Pt(3);p.paragraph_format.space_after=Pt(4)
 r=p.add_run('u модели(X*) = '+fnum(d['umodel'],7)+' г/г  ·  ПРЕДВАРИТЕЛЬНО');r.bold=True;r.font.color.rgb=RGBColor(150,63,47)
 p=doc.add_paragraph('Полная расширенная неопределённость U результата пробы не рассчитывалась: не закрыты данные весов, принятие регрессионной модели и сквозной бюджет пробы. ДИ на графике относится только к средней линии Y(X).')
 p.paragraph_format.space_after=Pt(3);p.runs[0].italic=True
 p=doc.add_paragraph('Методист: ____________________    Дата: ______________    Решение о допуске: НЕ ПОДТВЕРЖДЕНО')
 p.paragraph_format.space_before=Pt(3)
 for r in p.runs:r.font.size=Pt(7.5)
 filename=f'/mnt/data/EN13132_v15_Отчёт_{name}.docx';doc.save(filename)
 return filename

if __name__=='__main__':
 for symbol,eng in [('МТБЭ','MTBE'),('МТАЭ','MTAE')]:
  print(document(f'/mnt/data/EN13132_v15_{eng}.xlsx',symbol))
