#!/usr/bin/env python3
"""Validate and atomically replace questions.json from questions.xlsx without dependencies."""
import json,zipfile,xml.etree.ElementTree as ET,pathlib,sys
X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
file=pathlib.Path('questions.xlsx')
with zipfile.ZipFile(file) as z:
    strs=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        root=ET.fromstring(z.read('xl/sharedStrings.xml'))
        strs=[''.join(t.text or '' for t in si.iter(X+'t')) for si in root.findall(X+'si')]
    # worksheet from first tab; Excel templates place questions on the first sheet
    sheet=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    rows=[]
    for row in sheet.findall('.//'+X+'sheetData/'+X+'row'):
        values=['']*6
        for cell in row.findall(X+'c'):
            addr=cell.attrib['r']; col=''.join(c for c in addr if c.isalpha())
            idx=0
            for char in col:idx=idx*26+(ord(char)-64)
            if idx<1 or idx>6:continue
            typ=cell.attrib.get('t','')
            v=cell.find(X+'v')
            if typ=='inlineStr':value=''.join(t.text or '' for t in cell.iter(X+'t'))
            elif v is None:value=''
            elif typ=='s':value=strs[int(v.text or 0)]
            else:value=v.text or ''
            values[idx-1]=str(value).strip()
        if any(values):rows.append((row.attrib.get('r','?'),values))
if not rows or rows[0][1]!=['Въпрос','А','Б','В','Г','Верен']:
    sys.exit('Грешка: заглавията трябва да са Въпрос, А, Б, В, Г, Верен.')
questions=[]; errors=[]
for line,vals in rows[1:]:
    if any(not v for v in vals):errors.append('Ред '+line+': празна клетка.');continue
    answer=vals[5].upper().replace('A','А').replace('B','Б').replace('C','В').replace('D','Г')
    if answer not in 'АБВГ' or len(answer)!=1:errors.append('Ред '+line+': верният отговор трябва да е А, Б, В или Г.');continue
    questions.append({'q':vals[0],'a':vals[1:5],'ok':'АБВГ'.index(answer)})
if len(questions)<7:errors.append('Нужни са поне 7 въпроса за всичките мисии.')
if errors:sys.exit('НЕВАЛИДЕН EXCEL; старата база е запазена:\n'+'\n'.join(errors))
p=pathlib.Path('questions.json');p.write_text(json.dumps({'title':'Учебна мисия','questions':questions},ensure_ascii=False,indent=2)+'\n',encoding='utf8')
print('Успешно: заредени',len(questions),'въпроса. Старите въпроси са заменени.')
