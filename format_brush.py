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
        self.root.title("Word 表格格式刷")
        self.root.geometry("700x500")
        self.root.minsize(600, 400)
        
        self.style = ttk.Style("cosmo")
        
        main_frame = ttk.Frame(root, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="Word 表格格式刷", 
                 font=("Microsoft YaHei", 18, "bold")).pack(pady=(0,10))
        ttk.Label(main_frame, text="从模板文件复制表格格式到其他文件", 
                 font=("Microsoft YaHei", 10), bootstyle="secondary").pack(pady=(0,10))
        
        # 模板文件
        template_lf = ttk.Labelframe(main_frame, text=" 模板文件（格式正确的文件） ", bootstyle="info")
        template_lf.pack(fill=X, pady=(0,10))
        template_frame = ttk.Frame(template_lf, padding=10)
        template_frame.pack(fill=X)
        
        self.template_var = tk.StringVar(value="未选择文件")
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
        
    def select_template(self):
        file = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx")])
        if file:
            self.template_file = file
            self.template_var.set(os.path.basename(file))
            
    def add_targets(self):
        files = filedialog.askopenfilenames(filetypes=[("Word文件", "*.docx")])
        for f in files:
            if f not in self.target_files:
                self.target_files.append(f)
                self.target_listbox.insert(END, os.path.basename(f))
        self.count_var.set(f"共 {len(self.target_files)} 个文件")
        
    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            for f in glob.glob(os.path.join(folder, "**/*.docx"), recursive=True):
                if f not in self.target_files:
                    self.target_files.append(f)
                    self.target_listbox.insert(END, os.path.basename(f))
            self.count_var.set(f"共 {len(self.target_files)} 个文件")
            
    def clear_targets(self):
        self.target_files.clear()
        self.target_listbox.delete(0, END)
        self.count_var.set("共 0 个文件")
        
    def get_table_formats(self, docx_path):
        """从模板文件提取所有表格的格式信息"""
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml = z.read('word/document.xml')
        
        root = etree.fromstring(xml)
        formats = []
        
        for tbl in root.findall(f'.//{ns}tbl'):
            tbl_format = {'rows': []}
            
            # 表格属性
            tbl_pr = tbl.find(f'{ns}tblPr')
            if tbl_pr is not None:
                tbl_format['tblPr'] = etree.tostring(tbl_pr)
            
            # 表格网格
            tbl_grid = tbl.find(f'{ns}tblGrid')
            if tbl_grid is not None:
                tbl_format['tblGrid'] = etree.tostring(tbl_grid)
            
            # 行和单元格
            for tr in tbl.findall(f'.//{ns}tr'):
                row_format = {'cells': []}
                
                # 行属性
                tr_pr = tr.find(f'{ns}trPr')
                if tr_pr is not None:
                    row_format['trPr'] = etree.tostring(tr_pr)
                
                for tc in tr.findall(f'{ns}tc'):
                    cell_format = {}
                    
                    # 单元格属性
                    tc_pr = tc.find(f'{ns}tcPr')
                    if tc_pr is not None:
                        cell_format['tcPr'] = etree.tostring(tc_pr)
                    
                    # 单元格内容（保留格式信息）
                    paragraphs = tc.findall(f'.//{ns}p')
                    para_formats = []
                    for p in paragraphs:
                        p_pr = p.find(f'{ns}pPr')
                        p_format = {}
                        if p_pr is not None:
                            p_format['pPr'] = etree.tostring(p_pr)
                        
                        # run格式
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
        
    def apply_format_to_file(self, target_path, formats):
        """将格式应用到目标文件"""
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        with zipfile.ZipFile(target_path, 'r') as z:
            xml = z.read('word/document.xml')
        
        root = etree.fromstring(xml)
        target_tables = root.findall(f'.//{ns}tbl')
        
        for i, tbl in enumerate(target_tables):
            if i >= len(formats):
                break
            
            fmt = formats[i]
            
            # 应用表格属性
            if 'tblPr' in fmt:
                old_pr = tbl.find(f'{ns}tblPr')
                if old_pr is not None:
                    tbl.remove(old_pr)
                new_pr = etree.fromstring(fmt['tblPr'])
                tbl.insert(0, new_pr)
            
            # 应用网格
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
            
            # 应用行和单元格格式
            target_rows = tbl.findall(f'.//{ns}tr')
            for r_idx, tr in enumerate(target_rows):
                if r_idx >= len(fmt['rows']):
                    break
                
                row_fmt = fmt['rows'][r_idx]
                
                # 行属性
                if 'trPr' in row_fmt:
                    old_tr_pr = tr.find(f'{ns}trPr')
                    if old_tr_pr is not None:
                        tr.remove(old_tr_pr)
                    new_tr_pr = etree.fromstring(row_fmt['trPr'])
                    tr.insert(0, new_tr_pr)
                
                # 单元格格式
                target_cells = tr.findall(f'{ns}tc')
                for c_idx, tc in enumerate(target_cells):
                    if c_idx >= len(row_fmt['cells']):
                        break
                    
                    cell_fmt = row_fmt['cells'][c_idx]
                    
                    # 单元格属性
                    if 'tcPr' in cell_fmt:
                        old_tc_pr = tc.find(f'{ns}tcPr')
                        if old_tc_pr is not None:
                            tc.remove(old_tc_pr)
                        new_tc_pr = etree.fromstring(cell_fmt['tcPr'])
                        tc.insert(0, new_tc_pr)
                    
                    # 段落格式
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
                        
                        # run格式
                        target_runs = p.findall(f'.//{ns}r')
                        for r_idx2, r in enumerate(target_runs):
                            if r_idx2 < len(p_fmt.get('rPrs', [])):
                                old_r_pr = r.find(f'{ns}rPr')
                                if old_r_pr is not None:
                                    r.remove(old_r_pr)
                                new_r_pr = etree.fromstring(p_fmt['rPrs'][r_idx2])
                                r.insert(0, new_r_pr)
        
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        
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
        
        formats = self.get_table_formats(self.template_file)
        
        success = 0
        fail = 0
        
        for i, target in enumerate(self.target_files):
            self.status_var.set(f"处理 {i+1}/{len(self.target_files)}")
            self.root.update()
            
            try:
                new_xml = self.apply_format_to_file(target, formats)
                
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
            except Exception as e:
                print(f"处理 {target} 出错: {e}")
                fail += 1
                
        self.status_var.set("完成")
        messagebox.showinfo("完成", f"格式复制完成\n\n成功: {success}\n失败: {fail}")

def main():
    root = ttk.Window(title="Word 表格格式刷", themename="cosmo", size=(700, 500))
    app = FormatBrushTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()