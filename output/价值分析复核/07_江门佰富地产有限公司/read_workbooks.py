import json,zipfile,posixpath
from pathlib import Path
from xml.etree import ElementTree as E
ROOT=Path('/home/ubuntu/data/2026buliang')
OUT=ROOT/'output/价值分析复核/07_江门佰富地产有限公司'
N={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def read_xlsx(path):
 with zipfile.ZipFile(path) as z:
  strings=[]
  if 'xl/sharedStrings.xml' in z.namelist():
   strings=[''.join(si.itertext()) for si in E.fromstring(z.read('xl/sharedStrings.xml'))]
  rels={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
  sheets=[]
  for sh in E.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet',N):
   target=rels[sh.attrib['{'+N['r']+'}id']]
   target=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
   rows=[]
   for row in E.fromstring(z.read(target)).findall('s:sheetData/s:row',N):
    cells={}
    for c in row.findall('s:c',N):
     v=c.find('s:v',N); text=v.text if v is not None else None
     if c.get('t')=='s' and text is not None:text=strings[int(text)]
     if c.get('t')=='inlineStr':text=''.join(t.text or '' for t in c.findall('.//s:t',N))
     formula=c.find('s:f',N)
     if text is not None or formula is not None:cells[c.get('r')]={'value':text,'formula':formula.text if formula is not None else None}
    if cells:rows.append({'row':row.get('r'),'cells':cells})
   sheets.append({'sheet':sh.get('name'),'rows':rows})
 return sheets
if __name__=='__main__':
 for key,p in [('BK','2026buliang_现有资料整理_20261114/07_江门佰富地产有限公司/银行档案/07_江门佰富_基础信息提取_v1.xlsx'),('AS','2026buliang_现有资料整理_20261114/07_江门佰富地产有限公司/查册成果/结构包抵押物清单_07_江门佰富地产有限公司.xlsx')]:
  sheets=read_xlsx(ROOT/p)
  (OUT/('evidence_'+key+'07.json')).write_text(json.dumps({'source':p,'sheets':sheets},ensure_ascii=False,indent=2),encoding='utf-8')
  print(key,p)
  for sh in sheets:
   print('SHEET',sh['sheet'],'rows',len(sh['rows']))
   for row in sh['rows'] if key=='BK' else sh['rows'][:6]:
    print(' | '.join(c+':'+str(v['value']) for c,v in row['cells'].items()))
