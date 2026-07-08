from docx import Document
import os
import glob

def show_all_content(directory):
    """
    显示所有文档表格中的内容
    """
    docx_files = glob.glob(os.path.join(directory, '**/*.docx'), recursive=True)
    
    all_texts = set()
    
    for doc_path in docx_files:
        try:
            doc = Document(doc_path)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            all_texts.add(text)
        except Exception as e:
            print(f"读取 {doc_path} 时出错: {e}")
    
    # 按长度排序并显示
    sorted_texts = sorted(all_texts, key=len, reverse=True)
    
    print(f"共找到 {len(sorted_texts)} 个不同的文本内容:")
    print("=" * 80)
    
    for i, text in enumerate(sorted_texts[:50]):  # 只显示前50个
        print(f"{i+1:3d}. {text[:100]}")
    
    if len(sorted_texts) > 50:
        print(f"\n... 还有 {len(sorted_texts) - 50} 个文本未显示")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    show_all_content(base_dir)