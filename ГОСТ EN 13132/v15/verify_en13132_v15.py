#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Strict XML and numeric regression tests of the v15 presentation release.
Tests compare v14.3 base numerics vs v15, without changing either Excel.
"""
import re, zipfile, json, math, sys
from pathlib import Path
from lxml import etree as ET
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ROOT=Path('/mnt/data')

def excel(path):
 z=zipfile.ZipFile(path); w=ET.fromstring(z.read('xl/workbook.xml'))
 tabs=[(a.get('name'),a.get('state','visible')) for a in w.find('m:sheets',ns)]
 def cel(i,r):
  e=ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml')).find(f'.//m:c[@r="{r}"]',ns)
  if e is None:return None,None
  v=e.find('m:v',ns);f=e.find('m:f',ns)
  txt=v.text if v is not None else None
  try:txt=float(txt)
  except (TypeError,ValueError):pass
  return txt,''.join(f.itertext()) if f is not None else None
 charts=[a for a in z.namelist() if re.search(r'/(?:chart)\d+\.xml$',a)]
 return z,tabs,cel,charts

def run(src,out,kind):
 sz,_,s,_=excel(src); z,tabs,c,charts=excel(out)
 tests=[]
 def ck(name,passed,detail=''):tests.append(dict(test=name,pass_=bool(passed),detail=detail))
 ck('Exactly two user-visible worksheets', [x for x,st in tabs if st=='visible']==['01_Ввод','06_Отчёт'])
 ck('Every backend sheet retained',len(tabs)==14)
 ck('Charts retained',len(charts)==len([a for a in sz.namelist() if re.search(r'/chart\d+\.xml$',a)]))
 ck('P0 contracts retained',(c(12,'F5')[0] is not None) and c(12,'B27')[0]>0)
 for i in range(7,13):
  for col in ('B','C','D','E','F'):
   a=s(1,f'{col}{i}')[0];b=c(1,f'{col}{i}')[0]
   ck(f'Unchanged raw {col}{i}',a==b)
 for a in ('B11','B12','B21','B26','B31','B32','B33','B34','B35'):
  i=3;v,form=c(i,a);v0,_=s(i,a)
  if isinstance(v0,float):ck('Numeric invariant '+a,isinstance(v,float) and math.isclose(v,v0,abs_tol=5e-12,rel_tol=5e-10))
  else:ck('Invariant '+a,v==v0)
 for x in ('B19','B20','B21','B23','C25','B26'):
  v,_=c(6,x);base,_=s(6,x)
  if isinstance(base,float):ck('Budget same '+x,isinstance(v,float) and math.isclose(base,v,rel_tol=5e-10,abs_tol=5e-12))
 ck('Unsupported expanded U removed',c(6,'B27')[0]=='НЕ РАССЧИТАНА')
 ck('MC freshness visibly linked',"14_Контроль" in (c(6,'B28')[1] or ''))
 ck('Blank X* stays absent',c(2,'B27')[0] is None)
 ck('Blank individual masses u stay blank',all(c(1,f'T{i}')[0] is None and c(1,f'U{i}')[0] is None for i in range(7,13)))
 ck('Central model purity re-evaluates',isinstance(c(2,'B7')[0],float))
 ck('Numeric chart references retained',all(b'<c:xVal>' in z.read(a) or b':xVal>' in z.read(a) for a in charts if b'scatterChart' in z.read(a)))
 for i in (7,8,9):
  ck('Visible method control B'+str(i),c(1,'B'+str(28+i-7))[0] is not None)
 if kind=='BLANK':
  ck('Blank does not invent calibration readings', all(c(1,f'B{i}')[0] is None for i in range(7,13)))
  ck('Blank has no final u', c(6,'B26')[0] is None)
 fail=[t for t in tests if not t['pass_']]
 return dict(name=kind,passed=len(tests)-len(fail),total=len(tests),failed=fail)

if __name__=='__main__':
 output=[]
 for kind in ('MTBE','MTAE','LAB_BLANK'):
  k=kind if kind!='LAB_BLANK' else 'BLANK'
  d=run(ROOT/f'EN13132_v14_3_{k}.xlsx',ROOT/f'EN13132_v15_{kind}.xlsx',kind)
  print(kind,d['passed'],d['total'],d['failed'][:5]);output.append(d)
 (ROOT/'EN13132_v15_QA.json').write_text(json.dumps(output,ensure_ascii=False,indent=2),encoding='utf-8')
 sys.exit(0 if not any(x['failed'] for x in output) else 1)
