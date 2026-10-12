#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Final presentation layer for EN 13132 v14.3 without rewriting underlying calculation engine.
Visible worksheets: 01_Ввод, 06_Отчёт. All other sheets preserved and hidden (ordinary Excel Hidden).
Inputs are linked to existing source-of-truth cells; original raw calibration observations are untouched.
Usage: python finalize_en13132_v15.py --source ...xlsx --output ...xlsx
"""
import argparse, os, tempfile, zipfile
from lxml import etree as ET
from artifact_tool import Blob, SpreadsheetFile

FONT='#18324C'; DARK='#24486D'; HEAD='#DCE9F5'; INPUT='#EDF7EA'; SOFT='#F4F7FA'; TEXT='#3B4957'

PURE=[
 ('A28','Нижняя чистота аналита, %','B28',99,'02_Модель','B5'),
 ('A29','Верхняя чистота аналита, %','B29',100,'02_Модель','B6'),
 ('A30','Нижняя чистота ИС, %','B30',99,'02_Модель','B10'),
 ('A31','Верхняя чистота ИС, %','B31',100,'02_Модель','B11'),
 ('A32','Корреляция чистот ρ','B32',0,'02_Модель','B18'),
 ('A33','Δm (сценарий), мг','B33',0.1,'02_Модель','B15'),
 ('A34','Рабочая X* (необязательно)','B34',None,'02_Модель','B27'),
 ('A35','Обоснование ручного X*','B35',None,'02_Модель','B28'),
]
BAL=[
 ('D28','Модель весов','E28',None,'09_Весы','B6'),
 ('D29','Серийный номер','E29',None,'09_Весы','B7'),
 ('D30','Цена деления, г','E30',0.000001,'09_Весы','B8'),
 ('D31','Отсчётов на массу нетто','E31',2,'09_Весы','B9'),
 ('D32','u повторяемости, г','E32',None,'09_Весы','B10'),
 ('D33','U по свидетельству, г','E33',None,'09_Весы','B11'),
 ('D34','k по свидетельству','E34',None,'09_Весы','B12'),
 ('D35','Предел линейности, г','E35',None,'09_Весы','B13'),
 ('D36','Предел эксцентриситета, г','E36',None,'09_Весы','B14'),
 ('D37','Прочая u, г','E37',None,'09_Весы','B15'),
]
TRACE=[
 ('G28','Паспорт аналита, ID','H28',None,'02_Модель','B22'),
 ('G29','Паспорт ИС, ID','H29',None,'02_Модель','B23'),
 ('G30','Свидетельство весов, ID','H30',None,'09_Весы','B16'),
 ('G31','Схема определения массы','H31',None,'09_Весы','B23'),
 ('G32','Корреляция навесок','H32',None,'09_Весы','B24'),
]

def update_sheet(w, source, target_name, blank):
 inp=w.worksheets.get_item('01_Ввод'); rpt=w.worksheets.get_item('06_Отчёт')
 inp.get_range('A1').values=[['ГРАДУИРОВКА | '+target_name+' | ВВОД И ПРОСЛЕЖИВАЕМОСТЬ']]
 inp.get_range('A3').values=[['Заполните массы (г) и измеренные Y1–Y3. Xэфф рассчитывается автоматически; зелёные ячейки ниже заполняет методист.']]
 inp.get_range('A4').values=[['Новые значения экспериментальных откликов требуют пересчёта Q–Q через генератор v14.3 (контроль на отчёте).']]
 inp.get_range('A26:H26').format={'fill':DARK,'font':{'color':'#FFFFFF','bold':True},'row_height':25}
 inp.get_range('A26').values=[['ЧИСТОТА И РАБОЧАЯ ТОЧКА']]
 inp.get_range('D26').values=[['ВЕСЫ METTLER TOLEDO']]
 inp.get_range('G26').values=[['ДОКУМЕНТЫ']]
 if not blank:
  # Preserve already recorded method parameters when producing a new generation.
  existing={}
  for _,_,cell,default,backend,src_cell in PURE+BAL+TRACE:
   existing[cell]=w.worksheets.get_item(backend).get_range(src_cell).values[0][0]
 else:
  existing={}
 for label_cell,label,input_cell,default,backend,src_cell in PURE+BAL+TRACE:
  inp.get_range(label_cell).values=[[label]]
  v=existing.get(input_cell,default)
  if isinstance(v,str) and v.startswith('='):v=default
  inp.get_range(input_cell).values=[[v]]
  inp.get_range(input_cell).format={'fill':INPUT,'font':{'color':'#205C36'},'row_height':21}
  # Guard empty cells: a bare reference to a blank Excel cell evaluates to 0,
  # which would silently turn missing u(m) into a confirmed zero and X* into 0.
  numeric = input_cell.startswith(('B','E')) and input_cell not in ('B35','E28','E29')
  expr = (f"=IF(ISNUMBER('01_Ввод'!{input_cell}),'01_Ввод'!{input_cell},\"\")" if numeric
          else f"=IF(LEN('01_Ввод'!{input_cell})>0,'01_Ввод'!{input_cell},\"\")")
  w.worksheets.get_item(backend).get_range(src_cell).formulas=[[expr]]
 # A few informational lines, all critical source inputs are now in this one visible sheet.
 inp.get_range('A39').values=[['ИНДИВИДУАЛЬНЫЕ u(m), г (необязательно)']]
 inp.get_range('A40:C40').values=[['Уровень','u(m аналита), г','u(m ИС), г']]
 inp.get_range('A40:C40').format={'fill':HEAD,'font':{'bold':True,'color':FONT}}
 for k in range(12):
  r=41+k;sr=7+k
  inp.get_range(f'A{r}').values=[[k+1]]
  for c,sourcecol in [('B','T'),('C','U')]:
   val=None if blank else w.worksheets.get_item('01_Ввод').get_range(f'{sourcecol}{sr}').values[0][0]
   if isinstance(val,str) and val.startswith('='):val=None
   inp.get_range(f'{c}{r}').values=[[val]]
   inp.get_range(f'{c}{r}').format={'fill':INPUT}
   inp.get_range(f'{sourcecol}{sr}').formulas=[[f"=IF(ISNUMBER('01_Ввод'!{c}{r}),'01_Ввод'!{c}{r},\"\")"]]
 inp.get_range('D40:H40').values=[['ПРАВИЛА ДЛЯ БЮДЖЕТА','','','','']]
 inp.get_range('D41').values=[['Чистота 99–100%: допущение типа B, не паспортное u.']]
 inp.get_range('D42').values=[['Цена деления 0,000001 г не равна u(m).']]
 inp.get_range('D43').values=[['Без паспортов, калибровки весов и EIV бюджет предварителен.']]
 inp.get_range('D44').values=[['X* = среднее X минимизирует только u_reg (равноточный OLS).']]
 inp.get_range('D45').values=[['U полного анализа пробы здесь НЕ рассчитывается.']]
 inp.get_range('A26:A38').format.column_width=31
 inp.get_range('B26:B55').format.column_width=19
 inp.get_range('C26:C55').format.column_width=19
 inp.get_range('D26:D55').format.column_width=31
 inp.get_range('E26:E55').format.column_width=19
 inp.get_range('F26:F55').format.column_width=17
 inp.get_range('G26:G55').format.column_width=28
 inp.get_range('H26:H55').format.column_width=22
 inp.get_range('A40:C52').set_number_format('0.000000')
 inp.get_range('B28:B31').set_number_format('0.000')
 inp.get_range('B33').set_number_format('0.000')
 inp.get_range('E30').set_number_format('0.000000')
 try:inp.freeze_panes.freeze_rows(6)
 except Exception:pass
 # Output: no unsupported expanded U=2uc, nor an implication that this is a full uncertainty of a sample result.
 rpt.get_range('A1').values=[['ПРОТОКОЛ ГРАДУИРОВКИ | '+target_name]]
 rpt.get_range('A3').values=[['ГОСТ EN 13132-2012: относительный коэффициент f и диагностическая OLS представлены отдельно.']]
 rpt.get_range('A17').values=[['Статус параметров весов']]
 rpt.get_range('A18:D18').values=[['ИСТОЧНИК / ВКЛАД','u(X*), г/г','Дисперсия, (г/г)²','Доля, %']]
 rpt.get_range('A18:D18').format={'fill':DARK,'font':{'bold':True,'color':'#FFFFFF'}}
 rpt.get_range('A19').values=[['Регрессия (a,b и cov)']]
 rpt.get_range('A20').values=[['Чистота аналита']]
 rpt.get_range('A21').values=[['Чистота внутреннего стандарта']]
 rpt.get_range('A22').values=[['Корреляция чистот (со знаком)']]
 rpt.get_range('A23').values=[['Взвешивание навесок']]
 rpt.get_range('A25').values=[['Сумма дисперсий (условная модель)']]
 rpt.get_range('A26').values=[['u модели(X*), ПРЕДВАРИТЕЛЬНАЯ']]
 rpt.get_range('A27').values=[['U полного результата пробы']]
 rpt.get_range('B27').values=[['НЕ РАССЧИТАНА']]
 rpt.get_range('D27').values=[['Нет полного бюджета пробы и принятой модели']]
 rpt.get_range('A28').values=[['Актуальность графиков и Q–Q']]
 rpt.get_range('B28').formulas=[["='14_Контроль'!B15"]]
 
 rpt.get_range('A30').values=[['95%-ные полосы относятся к средней линии Y(X), а не к результату пробы.']]
 rpt.get_range('A31').values=[['Незакрытый бюджет и допуск: см. реестр BUD-01…BUD-20 (скрыт).']]
 # Compact final flags; keep one visible freshness warning.
 rpt.get_range('A32:D34').values=[[None]*4 for _ in range(3)]
 rpt.get_range('A35:C35').values=[['№','Xэфф, г/г','Yср, 1']]
 rpt.get_range('D35:F35').values=[['f=X/Y','X min','X max']]
 for rr, text in {6:'Допуск не подтверждён',7:'Отдельные инъекции',8:'OLS по средним',9:'Не отдельный вклад',10:'Описательный показатель',11:'Диагностика, не допуск',12:'LoF не выявлен ≠ линейность доказана',13:'См. F крит.',14:'При равных повторах',15:'Минимум только u_reg, при OLS',17:'Сценарий до подтверждения весов'}.items():
  rpt.get_range(f'D{rr}').values=[[text]]
 rpt.get_range('A5:D5').format={'fill':DARK,'font':{'color':'#FFFFFF','bold':True}}
 rpt.get_range('A19:D23').format.row_height=21
 rpt.get_range('B19:C26').set_number_format('0.000000000')
 rpt.get_range('D19:D23').set_number_format('0.00')
 rpt.get_range('A26:D27').format={'fill':SOFT,'font':{'bold':True,'color':FONT},'row_height':22}
 rpt.get_range('B27').format={'font':{'bold':True,'color':'#AD3838'}}
 rpt.get_range('A35:F35').format={'fill':HEAD,'font':{'bold':True,'color':FONT}}
 rpt.get_range('A1:A49').format.column_width=35
 rpt.get_range('B1:B49').format.column_width=20
 rpt.get_range('C1:C49').format.column_width=19
 rpt.get_range('D1:D49').format.column_width=37
 rpt.get_range('E35:F49').format.column_width=17
 rpt.get_range('A30:D31').format.wrap_text=True
 rpt.get_range('A30:D31').format.row_height=29
 return w

def hide_backend_xlsx(path):
 # OOXML workbook metadata has no public visibility setter in artifact_tool.
 # Patch ONLY Excel workbook metadata after artifact_tool has edited the actual sheets/formulas.
 ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; q=lambda n:'{'+ns+'}'+n
 fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=os.path.dirname(path));os.close(fd)
 with zipfile.ZipFile(path,'r') as src,zipfile.ZipFile(tmp,'w') as dst:
  for info in src.infolist():
   data=src.read(info.filename)
   if info.filename=='xl/workbook.xml':
    root=ET.fromstring(data)
    for sh in root.find(q('sheets')):
     if sh.get('name') in ('01_Ввод','06_Отчёт'):
      sh.attrib.pop('state',None)
     else:sh.set('state','hidden')
    bv=root.find(q('bookViews'))
    if bv is not None:
     v=bv.find(q('workbookView'))
     if v is not None:v.set('activeTab','0');v.set('firstSheet','0')
    data=ET.tostring(root,encoding='UTF-8',xml_declaration=True)
   dst.writestr(info,data)
 os.replace(tmp,path)

def build(source,output,mode):
 w=SpreadsheetFile.import_xlsx(Blob.load(source))
 update_sheet(w,source,mode,mode=='BLANK')
 SpreadsheetFile.export_xlsx(w).save(output)
 hide_backend_xlsx(output)
 return output

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--source',required=True);p.add_argument('--output',required=True);p.add_argument('--mode',required=True)
 a=p.parse_args();print(build(a.source,a.output,a.mode))
