# -*- coding: utf-8 -*-
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import glob
import zipfile
import shutil
from lxml import etree

class FormatBrushTool:
    def __init__(self, root):
        self.root = root
        self.root.title("文档/表格 格式刷")
        self.root.geometry("750x550")
        self.root.minsize(650, 450)
        
        self.style = ttk.Style("cosmo")
        
        main_frame = ttk.Frame(root, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        
        ttk.Label(main_frame, text="文档/表格 格式刷", 
                 font=("Microsoft YaHei", 18, "bold")).pack(pady=(0,5))
        ttk.Label(main_frame, text="支持 Word (.docx) 和 Excel (.xlsx)", 
                 font=("Microsoft YaHei", 10), bootstyle="secondary").pack(pady=(0,10))
        
        # 模板文件
        template_lf = ttk.Labelframe(main_frame, text=" 模板文件（格式正确的文件） ", bootstyle="info")
        template_lf.pack(fill=X, pady=(0,10))
        template_frame = ttk.Frame(template_lf, padding=10)
        template_frame.pack(fill=X)
        
        self.template_var = tk.StringVar(value="未选择文件")
        self.template_type = None
        ttk.Label(template_frame, textvariable=self.template_var, 
                 font=("Microsoft YaHei", 10), bootstyle="info").pack(side=LEFT, fill=X, expand=True)
        ttk.Button(template_frame, text="选择模板", command=self.select_template,
                  bootstyle="info-outline", width=12).pack(side=RIGHT)
        
        # 目标文件
        target_lf = ttk.Labelframe(main_frame, text=" 目标文件（需要修改格式的文件） ", bootstyle="success")
        target_lf.pack(fill=BOTH, expand=True, pady=(0,10))
        target_frame = ttk.Frame(target_lf, padding=10)
        target_frame.pack(fill=BOTH, expand=True)
        
        self.target_listbox = tk.Listbox(target_frame, height=6, selectmode=EXTENDED,
                                        font=("Consolas", 9))
        self.target_listbox.pack(fill=BOTH, expand=True)
        
        btn_frame = ttk.Frame(target_frame)
        btn_frame.pack(fill=X, pady=(8,0))
        ttk.Button(btn_frame, text="添加文件", command=self.add_targets,
                  bootstyle="outline", width=10).pack(side=LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="添加文件夹", command=self.add_folder,
                  bootstyle="outline", width=10).pack(side=LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="清空", command=self.clear_targets,
                  bootstyle="danger-outline", width=8).pack(side=LEFT)
        
        self.count_var = tk.StringVar(value="共 0 个文件")
        ttk.Label(btn_frame, textvariable=self.count_var,
                 font=("Microsoft YaHei", 9), bootstyle="secondary").pack(side=RIGHT)
        
        # 执行按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=X)
        ttk.Button(bottom_frame, text="执行格式复制", command=self.execute,
                  bootstyle="success", width=15).pack(side=LEFT)
        
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom_frame, textvariable=self.status_var,
                 font=("Microsoft YaHei", 9), bootstyle="secondary").pack(side=RIGHT)
        
        self.template_file = None
        self.target_files = []
        
    def _get_file_type(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == '.docx':
            return 'word'
        elif ext == '.xlsx':
            return 'excel'
        return None
        
    def select_template(self):
        file = filedialog.askopenfilename(filetypes=[
            ("文档/表格", "*.docx *.xlsx"),
            ("Word文件", "*.docx"),
            ("Excel文件", "*.xlsx")
        ])
        if file:
            self.template_file = file
            self.template_type = self._get_file_type(file)
            type_label = "Word" if self.template_type == 'word' else "Excel"
            self.template_var.set(f"{os.path.basename(file)} [{type_label}]")
            
    def add_targets(self):
        files = filedialog.askopenfilenames(filetypes=[
            ("文档/表格", "*.docx *.xlsx"),
            ("Word文件", "*.docx"),
            ("Excel文件", "*.xlsx")
        ])
        for f in files:
            if f not in self.target_files:
                self.target_files.append(f)
                ftype = self._get_file_type(f)
                label = "W" if ftype == 'word' else "E"
                self.target_listbox.insert(END, f"[{label}] {os.path.basename(f)}")
        self.count_var.set(f"共 {len(self.target_files)} 个文件")
        
    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            for ext in ['**/*.docx', '**/*.xlsx']:
                for f in glob.glob(os.path.join(folder, ext), recursive=True):
                    if f not in self.target_files:
                        self.target_files.append(f)
                        ftype = self._get_file_type(f)
                        label = "W" if ftype == 'word' else "E"
                        self.target_listbox.insert(END, f"[{label}] {os.path.basename(f)}")
            self.count_var.set(f"共 {len(self.target_files)} 个文件")
            
    def clear_targets(self):
        self.target_files.clear()
        self.target_listbox.delete(0, END)
        self.count_var.set("共 0 个文件")
        
    # ========== Word 格式处理 ==========
    def get_word_formats(self, docx_path):
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml = z.read('word/document.xml')
        
        root = etree.fromstring(xml)
        formats = []
        
        for tbl in root.findall(f'.//{ns}tbl'):
            tbl_format = {'rows': []}
            
            tbl_pr = tbl.find(f'{ns}tblPr')
            if tbl_pr is not None:
                tbl_format['tblPr'] = etree.tostring(tbl_pr)
            
            tbl_grid = tbl.find(f'{ns}tblGrid')
            if tbl_grid is not None:
                tbl_format['tblGrid'] = etree.tostring(tbl_grid)
            
            for tr in tbl.findall(f'.//{ns}tr'):
                row_format = {'cells': []}
                
                tr_pr = tr.find(f'{ns}trPr')
                if tr_pr is not None:
                    row_format['trPr'] = etree.tostring(tr_pr)
                
                for tc in tr.findall(f'{ns}tc'):
                    cell_format = {}
                    
                    tc_pr = tc.find(f'{ns}tcPr')
                    if tc_pr is not None:
                        cell_format['tcPr'] = etree.tostring(tc_pr)
                    
                    paragraphs = tc.findall(f'.//{ns}p')
                    para_formats = []
                    for p in paragraphs:
                        p_pr = p.find(f'{ns}pPr')
                        p_format = {}
                        if p_pr is not None:
                            p_format['pPr'] = etree.tostring(p_pr)
                        
                        runs = p.findall(f'.//{ns}r')
                        run_formats = []
                        for r in runs:
                            r_pr = r.find(f'{ns}rPr')
                            if r_pr is not None:
                                run_formats.append(etree.tostring(r_pr))
                        p_format['rPrs'] = run_formats
                        para_formats.append(p_format)
                    
                    cell_format['paragraphs'] = para_formats
                    row_format['cells'].append(cell_format)
                
                tbl_format['rows'].append(row_format)
            
            formats.append(tbl_format)
        
        return formats
        
    def apply_word_format(self, target_path, formats):
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        with zipfile.ZipFile(target_path, 'r') as z:
            xml = z.read('word/document.xml')
        
        root = etree.fromstring(xml)
        target_tables = root.findall(f'.//{ns}tbl')
        
        for i, tbl in enumerate(target_tables):
            if i >= len(formats):
                break
            
            fmt = formats[i]
            
            if 'tblPr' in fmt:
                old_pr = tbl.find(f'{ns}tblPr')
                if old_pr is not None:
                    tbl.remove(old_pr)
                new_pr = etree.fromstring(fmt['tblPr'])
                tbl.insert(0, new_pr)
            
            if 'tblGrid' in fmt:
                old_grid = tbl.find(f'{ns}tblGrid')
                if old_grid is not None:
                    tbl.remove(old_grid)
                new_grid = etree.fromstring(fmt['tblGrid'])
                tbl_pr = tbl.find(f'{ns}tblPr')
                if tbl_pr is not None:
                    tbl_pr.addnext(new_grid)
                else:
                    tbl.insert(0, new_grid)
            
            target_rows = tbl.findall(f'.//{ns}tr')
            for r_idx, tr in enumerate(target_rows):
                if r_idx >= len(fmt['rows']):
                    break
                
                row_fmt = fmt['rows'][r_idx]
                
                if 'trPr' in row_fmt:
                    old_tr_pr = tr.find(f'{ns}trPr')
                    if old_tr_pr is not None:
                        tr.remove(old_tr_pr)
                    new_tr_pr = etree.fromstring(row_fmt['trPr'])
                    tr.insert(0, new_tr_pr)
                
                target_cells = tr.findall(f'{ns}tc')
                for c_idx, tc in enumerate(target_cells):
                    if c_idx >= len(row_fmt['cells']):
                        break
                    
                    cell_fmt = row_fmt['cells'][c_idx]
                    
                    if 'tcPr' in cell_fmt:
                        old_tc_pr = tc.find(f'{ns}tcPr')
                        if old_tc_pr is not None:
                            tc.remove(old_tc_pr)
                        new_tc_pr = etree.fromstring(cell_fmt['tcPr'])
                        tc.insert(0, new_tc_pr)
                    
                    target_paras = tc.findall(f'.//{ns}p')
                    for p_idx, p in enumerate(target_paras):
                        if p_idx >= len(cell_fmt['paragraphs']):
                            break
                        
                        p_fmt = cell_fmt['paragraphs'][p_idx]
                        
                        if 'pPr' in p_fmt:
                            old_p_pr = p.find(f'{ns}pPr')
                            if old_p_pr is not None:
                                p.remove(old_p_pr)
                            new_p_pr = etree.fromstring(p_fmt['pPr'])
                            p.insert(0, new_p_pr)
                        
                        target_runs = p.findall(f'.//{ns}r')
                        for r_idx2, r in enumerate(target_runs):
                            if r_idx2 < len(p_fmt.get('rPrs', [])):
                                old_r_pr = r.find(f'{ns}rPr')
                                if old_r_pr is not None:
                                    r.remove(old_r_pr)
                                new_r_pr = etree.fromstring(p_fmt['rPrs'][r_idx2])
                                r.insert(0, new_r_pr)
        
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        
    # ========== Excel 格式处理 ==========
    def get_excel_formats(self, xlsx_path):
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        formats = {
            'styles_xml': None,
            'numFmts': [],
            'fonts': [],
            'fills': [],
            'borders': [],
            'cellXfs': [],
            'colWidths': [],
            'rowHeights': []
        }
        
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            # 提取样式
            if 'xl/styles.xml' in z.namelist():
                styles_xml = z.read('xl/styles.xml')
                styles_root = etree.fromstring(styles_xml)
                formats['styles_xml'] = styles_xml
                
                # 数字格式
                numFmts = styles_root.find('.//main:numFmts', ns)
                if numFmts is not None:
                    for nf in numFmts.findall('main:numFmt', ns):
                        formats['numFmts'].append(etree.tostring(nf))
                
                # 字体
                for font in styles_root.findall('.//main:fonts/main:font', ns):
                    formats['fonts'].append(etree.tostring(font))
                
                # 填充
                for fill in styles_root.findall('.//main:fills/main:fill', ns):
                    formats['fills'].append(etree.tostring(fill))
                
                # 边框
                for border in styles_root.findall('.//main:borders/main:border', ns):
                    formats['borders'].append(etree.tostring(border))
                
                # 单元格格式
                for xf in styles_root.findall('.//main:cellXfs/main:xf', ns):
                    formats['cellXfs'].append(etree.tostring(xf))
            
            # 提取每个 sheet 的列宽和行高
            for sheet_name in z.namelist():
                if sheet_name.startswith('xl/worksheets/sheet') and sheet_name.endswith('.xml'):
                    sheet_xml = z.read(sheet_name)
                    sheet_root = etree.fromstring(sheet_xml)
                    
                    # 列宽
                    cols = sheet_root.find('.//main:cols', ns)
                    if cols is not None:
                        for col in cols.findall('main:col', ns):
                            formats['colWidths'].append({
                                'min': col.get('min'),
                                'max': col.get('max'),
                                'width': col.get('width'),
                                'customWidth': col.get('customWidth')
                            })
                    
                    # 行高
                    for row in sheet_root.findall('.//main:sheetData/main:row', ns):
                        ht = row.get('ht')
                        if ht:
                            formats['rowHeights'].append({
                                'r': row.get('r'),
                                'ht': ht
                            })
        
        return formats
        
    def apply_excel_format(self, target_path, formats):
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        if formats['styles_xml'] is None:
            return None, None
        
        col_widths = None
        row_heights = None
        
        with zipfile.ZipFile(target_path, 'r') as z:
            # 处理样式
            if 'xl/styles.xml' in z.namelist():
                target_styles_xml = z.read('xl/styles.xml')
                target_root = etree.fromstring(target_styles_xml)
                
                # 替换数字格式
                target_numFmts = target_root.find('.//main:numFmts', ns)
                if target_numFmts is not None:
                    for child in list(target_numFmts):
                        target_numFmts.remove(child)
                elif formats['numFmts']:
                    target_numFmts = etree.SubElement(target_root.find('.//main:styleSheet', ns), '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}numFmts')
                
                for nf_xml in formats['numFmts']:
                    if target_numFmts is not None:
                        target_numFmts.append(etree.fromstring(nf_xml))
                
                # 替换字体
                target_fonts = target_root.find('.//main:fonts', ns)
                if target_fonts is not None:
                    for child in list(target_fonts):
                        target_fonts.remove(child)
                    for font_xml in formats['fonts']:
                        target_fonts.append(etree.fromstring(font_xml))
                
                # 替换填充
                target_fills = target_root.find('.//main:fills', ns)
                if target_fills is not None:
                    for child in list(target_fills):
                        target_fills.remove(child)
                    for fill_xml in formats['fills']:
                        target_fills.append(etree.fromstring(fill_xml))
                
                # 替换边框
                target_borders = target_root.find('.//main:borders', ns)
                if target_borders is not None:
                    for child in list(target_borders):
                        target_borders.remove(child)
                    for border_xml in formats['borders']:
                        target_borders.append(etree.fromstring(border_xml))
                
                # 替换单元格格式
                target_cellXfs = target_root.find('.//main:cellXfs', ns)
                if target_cellXfs is not None:
                    for child in list(target_cellXfs):
                        target_cellXfs.remove(child)
                    for xf_xml in formats['cellXfs']:
                        target_cellXfs.append(etree.fromstring(xf_xml))
                
                new_styles = etree.tostring(target_root, xml_declaration=True, encoding='UTF-8', standalone=True)
            else:
                new_styles = None
            
            # 处理列宽和行高
            for sheet_name in z.namelist():
                if sheet_name.startswith('xl/worksheets/sheet') and sheet_name.endswith('.xml'):
                    sheet_xml = z.read(sheet_name)
                    sheet_root = etree.fromstring(sheet_xml)
                    
                    # 应用列宽
                    if formats['colWidths']:
                        cols = sheet_root.find('.//main:cols', ns)
                        if cols is not None:
                            for child in list(cols):
                                cols.remove(child)
                        else:
                            sheetData = sheet_root.find('.//main:sheetData', ns)
                            if sheetData is not None:
                                cols = etree.Element('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cols')
                                sheetData.insert(0, cols)
                        
                        if cols is not None:
                            for cw in formats['colWidths']:
                                col_elem = etree.SubElement(cols, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}col')
                                col_elem.set('min', cw['min'])
                                col_elem.set('max', cw['max'])
                                col_elem.set('width', cw['width'])
                                if cw.get('customWidth'):
                                    col_elem.set('customWidth', cw['customWidth'])
                    
                    # 应用行高
                    if formats['rowHeights']:
                        for rh in formats['rowHeights']:
                            for row in sheet_root.findall('.//main:sheetData/main:row', ns):
                                if row.get('r') == rh['r']:
                                    row.set('ht', rh['ht'])
                                    row.set('customHeight', '1')
                                    break
                    
                    col_widths = etree.tostring(sheet_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    row_heights = col_widths
        
        return new_styles, col_widths
        
    # ========== 主逻辑 ==========
    def execute(self):
        if not self.template_file:
            messagebox.showwarning("提示", "请先选择模板文件")
            return
        if not self.target_files:
            messagebox.showwarning("提示", "请先添加目标文件")
            return
            
        msg = f"从模板「{os.path.basename(self.template_file)}」复制格式到 {len(self.target_files)} 个文件？"
        if not messagebox.askyesno("确认", msg):
            return
            
        self.status_var.set("提取模板格式...")
        self.root.update()
        
        # 根据模板类型提取格式
        if self.template_type == 'word':
            formats = self.get_word_formats(self.template_file)
        else:
            formats = self.get_excel_formats(self.template_file)
        
        success = 0
        fail = 0
        
        for i, target in enumerate(self.target_files):
            self.status_var.set(f"处理 {i+1}/{len(self.target_files)}")
            self.root.update()
            
            target_type = self._get_file_type(target)
            
            try:
                if target_type == 'word' and self.template_type == 'word':
                    new_xml = self.apply_word_format(target, formats)
                    temp = target + '.tmp'
                    with zipfile.ZipFile(target, 'r') as zin:
                        with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
                            for item in zin.infolist():
                                if item.filename == 'word/document.xml':
                                    zout.writestr(item, new_xml)
                                else:
                                    zout.writestr(item, zin.read(item.filename))
                    shutil.move(temp, target)
                    success += 1
                    
                elif target_type == 'excel' and self.template_type == 'excel':
                    new_styles, sheet_data = self.apply_excel_format(target, formats)
                    if new_styles:
                        temp = target + '.tmp'
                        with zipfile.ZipFile(target, 'r') as zin:
                            with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
                                for item in zin.infolist():
                                    if item.filename == 'xl/styles.xml':
                                        zout.writestr(item, new_styles)
                                    elif item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml') and sheet_data:
                                        zout.writestr(item, sheet_data)
                                    else:
                                        zout.writestr(item, zin.read(item.filename))
                        shutil.move(temp, target)
                        success += 1
                    else:
                        fail += 1
                        
                else:
                    fail += 1
                    
            except Exception as e:
                print(f"处理 {target} 出错: {e}")
                fail += 1
                
        self.status_var.set("完成")
        messagebox.showinfo("完成", f"格式复制完成\n\n成功: {success}\n失败: {fail}")

def main():
    root = ttk.Window(title="文档/表格 格式刷", themename="cosmo", size=(750, 550))
    app = FormatBrushTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
