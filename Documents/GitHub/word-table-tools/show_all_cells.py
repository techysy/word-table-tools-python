from docx import Document
import os

def show_all_cells(doc_path):
    """
    显示文档表格的所有单元格内容
    """
    doc = Document(doc_path)
    
    print(f"\n文件: {os.path.basename(doc_path)}")
    print("=" * 80)
    
    for table_idx, table in enumerate(doc.tables):
        print(f"\n表格 {table_idx}:")
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text:
                    print(f"  [{row_idx},{col_idx}]: {text}")

# 检查一个样本文件
sample_file = '公厕/中医院公厕.docx'
if os.path.exists(sample_file):
    show_all_cells(sample_file)