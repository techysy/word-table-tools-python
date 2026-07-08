from docx import Document
from docx.shared import Inches, Pt, Cm, Emu
import os
import glob

def adjust_table_width(doc_path, reduction_percent=10):
    """
    调整文档中表格的宽度
    
    参数:
    doc_path: 文档路径
    reduction_percent: 减少的百分比（默认10%）
    """
    try:
        doc = Document(doc_path)
        
        for table in doc.tables:
            # 获取当前表格宽度
            tbl = table._tbl
            tbl_pr = tbl.tblPr
            
            if tbl_pr is not None:
                tbl_w = tbl_pr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblW')
                if tbl_w is not None:
                    current_width = int(tbl_w.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w'))
                    new_width = int(current_width * (1 - reduction_percent / 100))
                    
                    # 设置新的表格宽度
                    tbl_w.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(new_width))
                    
                    # 调整列宽
                    tbl_grid = tbl.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblGrid')
                    if tbl_grid is not None:
                        grid_cols = tbl_grid.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridCol')
                        if grid_cols:
                            # 均匀调整每列宽度
                            col_width = new_width // len(grid_cols)
                            for grid_col in grid_cols:
                                grid_col.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(col_width))
                    
                    # 调整每个单元格的宽度
                    for row in table.rows:
                        for cell in row.cells:
                            tc = cell._tc
                            tc_pr = tc.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr')
                            if tc_pr is not None:
                                tc_w = tc_pr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcW')
                                if tc_w is not None:
                                    current_cell_width = int(tc_w.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w'))
                                    new_cell_width = int(current_cell_width * (1 - reduction_percent / 100))
                                    tc_w.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(new_cell_width))
        
        # 保存文档
        doc.save(doc_path)
        return True, f"成功调整表格宽度，减少{reduction_percent}%"
        
    except Exception as e:
        return False, f"错误: {e}"

def batch_adjust_tables(directory, reduction_percent=10):
    """
    批量调整目录下所有docx文档的表格宽度
    """
    results = {}
    
    # 查找所有docx文件
    docx_files = glob.glob(os.path.join(directory, '**/*.docx'), recursive=True)
    
    for doc_path in docx_files:
        rel_path = os.path.relpath(doc_path, directory)
        success, message = adjust_table_width(doc_path, reduction_percent)
        results[rel_path] = {'success': success, 'message': message}
    
    return results

if __name__ == "__main__":
    # 设置目录和调整参数
    base_dir = os.path.dirname(os.path.abspath(__file__))
    reduction_percent = 10  # 减少10%
    
    print(f"开始批量调整表格宽度，减少{reduction_percent}%...")
    results = batch_adjust_tables(base_dir, reduction_percent)
    
    # 统计结果
    success_count = sum(1 for r in results.values() if r['success'])
    fail_count = len(results) - success_count
    
    print(f"\n调整完成:")
    print(f"  总文件数: {len(results)}")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")
    
    if fail_count > 0:
        print(f"\n失败的文件:")
        for file_path, result in results.items():
            if not result['success']:
                print(f"  {file_path}: {result['message']}")