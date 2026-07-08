import os
from docx import Document

def analyze_docx(doc_path):
    try:
        doc = Document(doc_path)
        tables_info = []
        
        for i, table in enumerate(doc.tables):
            table_info = {
                'index': i,
                'rows': len(table.rows),
                'cols': len(table.columns),
                'cells': []
            }
            
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if text:
                        table_info['cells'].append({
                            'row': row_idx,
                            'col': col_idx,
                            'text': text[:100]  # 只取前100个字符
                        })
            
            tables_info.append(table_info)
        
        return tables_info
    except Exception as e:
        return f"错误: {e}"

def scan_directory(directory):
    results = {}
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, directory)
                results[rel_path] = analyze_docx(file_path)
    return results

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    results = scan_directory(base_dir)
    
    print(f"扫描完成，共找到 {len(results)} 个docx文件")
    print("\n表格结构分析:")
    
    for file_path, info in results.items():
        if isinstance(info, list):
            print(f"\n{file_path}:")
            for table in info:
                print(f"  表格{table['index']}: {table['rows']}行 x {table['cols']}列")
                if table['cells']:
                    print(f"    第一个单元格内容: {table['cells'][0]['text'][:80]}...")
        else:
            print(f"\n{file_path}: {info}")