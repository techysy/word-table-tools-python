# -*- coding: utf-8 -*-
import os
import sys

# 设置控制台编码
os.environ['PYTHONIOENCODING'] = 'utf-8'

from docx import Document

# 使用绝对路径
base_dir = r'C:\Users\ZJZHZF\Documents\指挥中心\5-1-1'
file_path = os.path.join(base_dir, '公厕', '中医院公厕.docx')

print(f"文件路径: {file_path}")
print(f"文件存在: {os.path.exists(file_path)}")

doc = Document(file_path)
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