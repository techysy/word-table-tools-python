# -*- coding: utf-8 -*-
from docx import Document

doc = Document('公厕/中医院公厕.docx')
body = doc.element.body
ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
tables = body.findall(f'.//{ns}tbl')
table = tables[0]
rows = table.findall(f'.//{ns}tr')

for i, row in enumerate(rows):
    cells = row.findall(f'{ns}tc')
    texts = []
    for cell in cells:
        ts = cell.findall(f'.//{ns}t')
        text = ''.join([t.text for t in ts if t.text])
        texts.append(text)
    print(f"Row {i}: {texts}")