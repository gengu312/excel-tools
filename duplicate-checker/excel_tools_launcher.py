#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具集合 - 启动器
每个功能在独立窗口中运行
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from collections import defaultdict
import re

def normalize_text(text):
    """标准化文本，移除所有空白字符用于比较"""
    if pd.isna(text):
        return ""
    
    # 将所有空白字符（空格、制表符、换行符等）移除
    normalized = re.sub(r'\s+', '', str(text))
    return normalized

def visualize_whitespace(text):
    """可视化空白字符"""
    if pd.isna(text):
        return ""
    
    # 将空格替换为·，制表符替换为→
    visual = str(text).replace(' ', '·').replace('\t', '→').replace('\n', '↵')
    return visual

class DuplicateCheckerWindow:
    """重复数据检测窗口"""
    def __init__(self, parent=None):
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("重复数据检测")
        self.window.geometry("900x700")
        
        # 数据存储
        self.df = None
        self.file_path = None
        
        # 创建界面
        self.create_interface()
        
        # 如果是独立运行，设置关闭事件
        if not parent:
            self.window.protocol("WM_DELETE_WINDOW", self.window.quit)
    
    def create_interface(self):
        """创建重复数据检测界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
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
                messagebox.showerror("错误", f"文件加载失败：{str(e)}")
    
    def check_duplicates(self):
        """检测重复数据"""
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
            
            # 创建标准化文本到原始文本的映射
            normalized_groups = defaultdict(list)
            
            for idx, original_text in column_data.items():
                if pd.isna(original_text) or original_text == 'nan':
                    continue
                
                normalized = normalize_text(original_text)
                normalized_groups[normalized].append((original_text, idx + 2))  # Excel行号从2开始
            
            # 找出有多个不同原始值的标准化组（即存在空格差异的重复数据）
            duplicates_found = False
            total_groups = 0
            total_items = 0
            
            for normalized, items in normalized_groups.items():
                if len(items) > 1:
                    # 检查是否真的有不同的原始值
                    unique_originals = list(set([item[0] for item in items]))
                    if len(unique_originals) > 1:
                        duplicates_found = True
                        total_groups += 1
                        total_items += len(items)
                        
                        # 按原始值分组显示
                        original_groups = defaultdict(list)
                        for original, row_num in items:
                            original_groups[original].append(row_num)
                        
                        for original, row_numbers in original_groups.items():
                            count = len(row_numbers)
                            row_numbers_str = ', '.join(map(str, row_numbers))
                            
                            # 可视化空白字符
                            original_display = visualize_whitespace(original)
                            normalized_display = visualize_whitespace(normalized) if normalized else "(空)"
                            
                            self.tree.insert('', 'end', values=(
                                original_display,
                                normalized_display,
                                count,
                                row_numbers_str
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

class DataGrouperWindow:
    """数据分组分析窗口"""
    def __init__(self, parent=None):
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("数据分组分析")
        self.window.geometry("900x700")
        
        # 数据存储
        self.df = None
        self.file_path = None
        
        # 创建界面
        self.create_interface()
        
        # 如果是独立运行，设置关闭事件
        if not parent:
            self.window.protocol("WM_DELETE_WINDOW", self.window.quit)
    
    def create_interface(self):
        """创建数据分组分析界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="选择Excel文件", command=self.select_file).grid(row=0, column=0, padx=(0, 10))
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.grid(row=0, column=1, sticky=tk.W)
        
        # 搜索筛选区域
        search_frame = ttk.LabelFrame(main_frame, text="模糊搜索筛选", padding="10")
        search_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(search_frame, text="搜索列：").grid(row=0, column=0, padx=(0, 10))
        self.search_column_var = tk.StringVar()
        self.search_column_combo = ttk.Combobox(search_frame, textvariable=self.search_column_var, state="readonly")
        self.search_column_combo.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Label(search_frame, text="搜索关键词：").grid(row=0, column=2, padx=(0, 10))
        self.search_keyword_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_keyword_var, width=20)
        self.search_entry.grid(row=0, column=3, padx=(0, 10))
        
        # 分析设置区域
        analysis_frame = ttk.LabelFrame(main_frame, text="分析设置", padding="10")
        analysis_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(analysis_frame, text="分析列：").grid(row=0, column=0, padx=(0, 10))
        self.analysis_column_var = tk.StringVar()
        self.analysis_column_combo = ttk.Combobox(analysis_frame, textvariable=self.analysis_column_var, state="readonly")
        self.analysis_column_combo.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(analysis_frame, text="开始分析", command=self.analyze_data_groups).grid(row=0, column=2, padx=(10, 0))
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="分析结果", padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 创建Treeview来显示结果
        columns = ('数据种类', '状态', '出现次数', '行号')
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
        button_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="导出结果", command=self.export_results).grid(row=0, column=0, padx=(0, 10))
        
        # 配置网格权重
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
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
                self.search_column_combo['values'] = columns
                if columns:
                    self.search_column_combo.set(columns[0])
                
                self.analysis_column_combo['values'] = columns
                if columns:
                    self.analysis_column_combo.set(columns[0])
                
                messagebox.showinfo("成功", f"文件加载成功！\n共 {len(self.df)} 行数据")
                
            except FileNotFoundError:
                messagebox.showerror("错误", "文件未找到，请检查文件路径是否正确。")
            except PermissionError:
                messagebox.showerror("错误", "文件被其他程序占用，请关闭Excel后重试。")
            except Exception as e:
                messagebox.showerror("错误", f"文件加载失败：{str(e)}")
    
    def analyze_data_groups(self):
        """分析数据分组（提取最后一个下划线后的数据种类，标注重复项）"""
        if self.df is None:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        if not self.search_column_var.get() or not self.analysis_column_var.get():
            messagebox.showwarning("警告", "请选择搜索列和分析列")
            return
        
        if not self.search_keyword_var.get():
            messagebox.showwarning("警告", "请输入搜索关键词")
            return
        
        try:
            search_column = self.search_column_var.get()
            analysis_column = self.analysis_column_var.get()
            keyword = self.search_keyword_var.get()
            
            # 清空之前的结果
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 模糊搜索筛选数据
            filtered_df = self.df[self.df[search_column].astype(str).str.contains(keyword, case=False, na=False)]
            
            if filtered_df.empty:
                messagebox.showinfo("结果", f"未找到包含关键词 '{keyword}' 的数据")
                return
            
            # 分析指定列的数据，提取最后一个下划线后的内容
            analysis_data = filtered_df[analysis_column].astype(str)
            
            # 统计最后一个下划线后的数据种类
            suffix_count = {}
            suffix_rows = {}
            
            for idx, value in analysis_data.items():
                if pd.isna(value) or value == 'nan':
                    continue
                
                # 查找最后一个下划线位置
                last_underscore = value.rfind('_')
                
                if last_underscore != -1:  # 找到下划线
                    # 最后一个下划线后的部分
                    suffix = value[last_underscore+1:]
                    
                    if suffix not in suffix_count:
                        suffix_count[suffix] = 0
                        suffix_rows[suffix] = []
                    
                    suffix_count[suffix] += 1
                    suffix_rows[suffix].append(idx + 2)  # Excel行号从2开始
            
            if not suffix_count:
                messagebox.showinfo("结果", "未找到包含下划线的数据")
                return
            
            # 显示结果，按出现次数排序
            sorted_suffixes = sorted(suffix_count.items(), key=lambda x: x[1], reverse=True)
            
            duplicate_count = 0
            unique_suffixes = set()
            
            for suffix, count in sorted_suffixes:
                row_numbers = ', '.join(map(str, suffix_rows[suffix]))
                is_duplicate = count > 1
                
                if is_duplicate:
                    duplicate_count += 1
                    status = "重复"
                else:
                    status = "唯一"
                
                unique_suffixes.add(suffix)
                
                self.tree.insert('', 'end', values=(
                    suffix,
                    status,
                    count,
                    row_numbers
                ))
            
            # 生成不重复清单
            unique_list = sorted(list(unique_suffixes))
            unique_list_str = ', '.join(unique_list)
            
            messagebox.showinfo(
                "分析结果", 
                f"分析完成！\n\n搜索关键词：{keyword}\n筛选出 {len(filtered_df)} 条数据\n\n数据种类统计：\n- 总种类数：{len(unique_suffixes)}\n- 重复种类数：{duplicate_count}\n- 唯一种类数：{len(unique_suffixes) - duplicate_count}\n\n不重复清单：\n{unique_list_str}"
            )
            
        except Exception as e:
            messagebox.showerror("错误", f"分析过程中出现错误：{str(e)}")
    
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
                        '数据种类': values[0],
                        '状态': values[1],
                        '出现次数': values[2],
                        '行号': values[3]
                    })
                
                # 创建DataFrame并保存
                result_df = pd.DataFrame(results)
                result_df.to_excel(save_path, index=False)
                
                messagebox.showinfo("成功", f"结果已导出到：{save_path}")
                
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

class ExcelToolsLauncher:
    """主启动器窗口"""
    def __init__(self, root):
        self.root = root
        self.root.title("Excel工具集合 - 启动器")
        self.root.geometry("400x350")
        
        # 设置窗口居中
        self.center_window()
        
        # 创建主界面
        self.create_main_interface()
    
    def center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_main_interface(self):
        """创建主界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="30")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel工具集合", font=("Arial", 20, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 副标题
        subtitle_label = ttk.Label(main_frame, text="选择需要使用的功能", font=("Arial", 12))
        subtitle_label.grid(row=1, column=0, pady=(0, 30))
        
        # 功能按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=20)
        
        # 功能按钮
        duplicate_btn = ttk.Button(
            button_frame, 
            text="重复数据检测", 
            command=self.open_duplicate_checker, 
            width=25
        )
        duplicate_btn.pack(pady=12)
        
        grouper_btn = ttk.Button(
            button_frame, 
            text="数据分组分析", 
            command=self.open_data_grouper, 
            width=25
        )
        grouper_btn.pack(pady=12)
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=20)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=4, column=0)
        
        about_btn = ttk.Button(bottom_frame, text="关于", command=self.show_about, width=12)
        about_btn.pack(side=tk.LEFT, padx=5)
        
        exit_btn = ttk.Button(bottom_frame, text="退出", command=self.root.quit, width=12)
        exit_btn.pack(side=tk.LEFT, padx=5)
    
    def open_duplicate_checker(self):
        """打开重复数据检测窗口"""
        try:
            DuplicateCheckerWindow(self.root)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开重复数据检测窗口：{str(e)}")
    
    def open_data_grouper(self):
        """打开数据分组分析窗口"""
        try:
            DataGrouperWindow(self.root)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开数据分组分析窗口：{str(e)}")
    
    def show_about(self):
        """显示关于信息"""
        about_text = (
            "Excel工具集合 v2.0\n\n"
            "功能介绍：\n"
            "• 重复数据检测\n"
            "  检测包含空格差异的重复数据\n\n"
            "• 数据分组分析\n"
            "  模糊搜索并分析数据分类\n\n"
            "特点：\n"
            "• 每个功能在独立窗口中运行\n"
            "• 支持.xls和.xlsx格式\n"
            "• 可视化空白字符差异\n\n"
            "作者：AI Assistant"
        )
        messagebox.showinfo("关于", about_text)

def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelToolsLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()