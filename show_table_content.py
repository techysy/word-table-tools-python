from docx import Document
import os

def show_table_content(doc_path):
    """
    显示文档表格的详细内容
    """
    doc = Document(doc_path)
    
    print(f"\n文件: {os.path.basename(doc_path)}")
    print("=" * 80)
    
    for i, table in enumerate(doc.tables):
        print(f"\n表格 {i}:")
        print(f"  行数: {len(table.rows)}")
        print(f"  列数: {len(table.columns)}")
        
        # 打印表格内容
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text:
                    print(f"  [{row_idx},{col_idx}]: {text[:150]}")

# 检查几个样本文件
sample_files = [
    '公厕/中医院公厕.docx',
    '垃圾站/西水岸垃圾站.docx',
    '乡镇中转站/龙台镇双龙垃圾压缩中转站.docx'
]

for file_path in sample_files:
    if os.path.exists(file_path):
        show_table_content(file_path)