from docx import Document
import xml.etree.ElementTree as ET

doc_path = '公厕/中医院公厕.docx'
doc = Document(doc_path)

table = doc.tables[0]
tbl = table._tbl

# 打印表格的XML
print("表格XML:")
print(ET.tostring(tbl, encoding='unicode'))

# 检查表格的所有行和列
print("\n表格结构:")
for row_idx, row in enumerate(table.rows):
    print(f"行 {row_idx}:")
    for col_idx, cell in enumerate(row.cells):
        print(f"  列 {col_idx}: {cell.text[:50]}")