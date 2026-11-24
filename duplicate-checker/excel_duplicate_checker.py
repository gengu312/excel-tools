#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel重复数据检测工具
功能：检测Excel表格中某一列存在的相同但包含多余空格的数据
作者：AI Assistant
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class ExcelDuplicateChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel重复数据检测工具")
        self.root.geometry("800x600")
        
        # 数据存储
        self.df = None
        self.file_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="选择Excel文件", command=self.select_file).grid(row=0, column=0, padx=(0, 10))
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.grid(row=0, column=1, sticky=tk.W)
        
        # 列选择区域
        column_frame = ttk.LabelFrame(main_frame, text="列选择", padding="10")
        column_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(column_frame, text="选择要检测的列：").grid(row=0, column=0, padx=(0, 10))
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(column_frame, textvariable=self.column_var, state="readonly")
        self.column_combo.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(column_frame, text="开始检测", command=self.check_duplicates).grid(row=0, column=2)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="检测结果", padding="10")
        result_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 创建Treeview来显示结果
        columns = ('原始值', '清理后值', '出现次数', '行号')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=15)
        
        # 设置列标题
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="导出结果", command=self.export_results).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="清理数据", command=self.clean_data).grid(row=0, column=1)
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
    
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # 检查文件扩展名并选择合适的引擎
                file_extension = os.path.splitext(file_path)[1].lower()
                
                if file_extension == '.xls':
                    # 对于.xls文件，尝试使用xlrd引擎
                    try:
                        self.df = pd.read_excel(file_path, engine='xlrd')
                    except ImportError:
                        # 如果xlrd未安装，提示用户安装
                        messagebox.showerror(
                            "缺少依赖", 
                            "读取.xls文件需要安装xlrd库。\n\n请在命令行中运行：\npip install xlrd\n\n或者将文件另存为.xlsx格式。"
                        )
                        return
                    except Exception as e:
                        # 如果xlrd读取失败，尝试openpyxl
                        try:
                            self.df = pd.read_excel(file_path, engine='openpyxl')
                        except Exception as e2:
                            messagebox.showerror(
                                "文件读取失败", 
                                f"无法读取.xls文件。\n\n错误信息：{str(e)}\n\n建议：\n1. 安装xlrd库：pip install xlrd\n2. 或将文件另存为.xlsx格式"
                            )
                            return
                elif file_extension == '.xlsx':
                    # 对于.xlsx文件，使用openpyxl引擎
                    self.df = pd.read_excel(file_path, engine='openpyxl')
                else:
                    # 默认情况，让pandas自动选择引擎
                    self.df = pd.read_excel(file_path)
                
                self.file_path = file_path
                
                # 更新文件标签
                filename = os.path.basename(file_path)
                self.file_label.config(text=f"已选择: {filename}")
                
                # 更新列选择下拉框
                columns = list(self.df.columns)
                self.column_combo['values'] = columns
                if columns:
                    self.column_combo.set(columns[0])
                
                messagebox.showinfo("成功", f"文件加载成功！\n共 {len(self.df)} 行数据")
                
            except FileNotFoundError:
                messagebox.showerror("错误", "文件未找到，请检查文件路径是否正确。")
            except PermissionError:
                messagebox.showerror("错误", "文件被其他程序占用，请关闭Excel后重试。")
            except Exception as e:
                error_msg = str(e)
                if "xlrd" in error_msg.lower():
                    messagebox.showerror(
                        "依赖库问题", 
                        "读取.xls文件需要xlrd库。\n\n解决方案：\n1. 安装xlrd：pip install xlrd\n2. 或将文件另存为.xlsx格式"
                    )
                else:
                    messagebox.showerror("错误", f"文件加载失败：{error_msg}\n\n可能的原因：\n1. 文件格式不正确\n2. 文件已损坏\n3. 缺少必要的依赖库")
    
    def normalize_text(self, text):
        """标准化文本，完全去除所有空白字符进行比较"""
        import re
        # 完全去除所有类型的空白字符（空格、制表符、换行符等）
        normalized = re.sub(r'\s+', '', str(text))
        return normalized
    
    def check_duplicates(self):
        """检测重复数据（包括各种空格差异）"""
        if self.df is None:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        if not self.column_var.get():
            messagebox.showwarning("警告", "请选择要检测的列")
            return
        
        try:
            column_name = self.column_var.get()
            
            # 清空之前的结果
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 获取指定列的数据
            column_data = self.df[column_name].astype(str)
            
            # 创建一个字典来存储标准化后的值和原始值的映射
            normalized_to_original = {}
            
            for idx, value in enumerate(column_data):
                if pd.isna(value) or value == 'nan':
                    continue
                    
                # 使用改进的标准化方法
                normalized_value = self.normalize_text(value)
                
                if normalized_value not in normalized_to_original:
                    normalized_to_original[normalized_value] = []
                
                normalized_to_original[normalized_value].append({
                    'original': value,
                    'row': idx + 2  # Excel行号从2开始（考虑标题行）
                })
            
            # 找出有多个不同原始值的标准化值
            duplicates_found = False
            total_groups = 0
            total_items = 0
            
            for normalized_value, original_list in normalized_to_original.items():
                # 获取所有不同的原始值
                unique_originals = list(set([item['original'] for item in original_list]))
                
                # 如果有多个不同的原始值，说明存在空格差异
                if len(unique_originals) > 1:
                    duplicates_found = True
                    total_groups += 1
                    
                    # 按原始值分组显示
                    for original in unique_originals:
                        # 计算该原始值出现的次数
                        count = sum(1 for item in original_list if item['original'] == original)
                        total_items += count
                        
                        # 获取所有行号
                        rows = [str(item['row']) for item in original_list if item['original'] == original]
                        row_numbers = ', '.join(rows)
                        
                        # 显示空格差异（用特殊符号标记空格）
                        original_display = original.replace(' ', '·').replace('\t', '→')
                        normalized_display = normalized_value.replace(' ', '·')
                        
                        # 插入到结果树中
                        self.tree.insert('', 'end', values=(
                            original_display,
                            normalized_display,
                            count,
                            row_numbers
                        ))
            
            if not duplicates_found:
                messagebox.showinfo("结果", "未发现存在空格差异的重复数据")
            else:
                messagebox.showinfo(
                    "检测结果", 
                    f"检测完成！\n\n发现 {total_groups} 组存在空格差异的重复数据\n共涉及 {total_items} 条记录\n\n说明：\n· 表示空格\n→ 表示制表符"
                )
                
        except Exception as e:
            messagebox.showerror("错误", f"检测过程中出现错误：{str(e)}")
    
    def export_results(self):
        """导出检测结果"""
        if not self.tree.get_children():
            messagebox.showwarning("警告", "没有可导出的结果")
            return
        
        try:
            # 选择保存路径
            save_path = filedialog.asksaveasfilename(
                title="保存检测结果",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if save_path:
                # 收集结果数据
                results = []
                for item in self.tree.get_children():
                    values = self.tree.item(item)['values']
                    results.append({
                        '原始值': values[0],
                        '清理后值': values[1],
                        '出现次数': values[2],
                        '行号': values[3]
                    })
                
                # 创建DataFrame并保存
                result_df = pd.DataFrame(results)
                result_df.to_excel(save_path, index=False)
                
                messagebox.showinfo("成功", f"结果已导出到：{save_path}")
                
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
    
    def clean_data(self):
        """清理数据中的多余空格"""
        if self.df is None:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        if not self.column_var.get():
            messagebox.showwarning("警告", "请选择要清理的列")
            return
        
        try:
            # 确认操作
            result = messagebox.askyesno(
                "确认", 
                "这将清理选定列中所有数据的首尾空格。\n是否继续？"
            )
            
            if result:
                column_name = self.column_var.get()
                
                # 清理数据
                self.df[column_name] = self.df[column_name].astype(str).str.strip()
                
                # 选择保存路径
                save_path = filedialog.asksaveasfilename(
                    title="保存清理后的文件",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
                
                if save_path:
                    self.df.to_excel(save_path, index=False)
                    messagebox.showinfo("成功", f"清理后的文件已保存到：{save_path}")
                
        except Exception as e:
            messagebox.showerror("错误", f"清理失败：{str(e)}")

def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelDuplicateChecker(root)
    root.mainloop()

if __name__ == "__main__":
    main()