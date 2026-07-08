from docx import Document
import xml.etree.ElementTree as ET
import os

def parse_table_xml(doc_path):
    """
    解析文档表格的XML结构
    """
    doc = Document(doc_path)
    
    print(f"\n文件: {os.path.basename(doc_path)}")
    print("=" * 80)
    
    # 获取文档的body
    body = doc.element.body
    
    # 查找所有表格
    tables = body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')
    
    for i, table in enumerate(tables):
        print(f"\n表格 {i}:")
        
        # 查找所有行
        rows = table.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        print(f"  行数: {len(rows)}")
        
        # 解析每一行
        for row_idx, row in enumerate(rows):
            # 查找所有单元格
            cells = row.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            
            # 检查行属性
            tr_pr = row.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trPr')
            if tr_pr is not None:
                # 检查行高
                tr_height = tr_pr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trHeight')
                if tr_height is not None:
                    height_val = tr_height.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    print(f"  行 {row_idx} 高度: {height_val}")
            
            # 解析每个单元格
            for cell_idx, cell in enumerate(cells):
                # 获取单元格文本
                paragraphs = cell.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                text = ""
                for p in paragraphs:
                    runs = p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    for run in runs:
                        if run.text:
                            text += run.text
                
                # 获取单元格宽度
                tc_pr = cell.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr')
                width = "未知"
                if tc_pr is not None:
                    tc_w = tc_pr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcW')
                    if tc_w is not None:
                        width = tc_w.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w')
                
                if text.strip():
                    print(f"    [{row_idx},{cell_idx}] (宽度:{width}): {text[:100]}")

# 检查几个样本文件
sample_files = [
    '公厕/中医院公厕.docx',
    '垃圾站/西水岸垃圾站.docx',
    '乡镇中转站/龙台镇双龙垃圾压缩中转站.docx'
]

for file_path in sample_files:
    if os.path.exists(file_path):
        parse_table_xml(file_path)