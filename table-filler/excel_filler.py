#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格填写工具主程序
用于将发货明细表数据按模板表格式进行整理和填写
"""

from os.path import basename, join, dirname, normpath
from os import startfile
import sys
import json
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
import threading
import os
import sys
from pathlib import Path

class ExcelColumnMapper:
    """列映射配置管理类"""
    
    def __init__(self):
        self.mappings = []
    
    def add_mapping(self, source_col, target_cols):
        """添加列映射关系
        
        Args:
            source_col (str): 发货明细表的列名
            target_cols (list): 模板表的列名列表（支持1对1或1对2）
        """
        self.mappings.append({
            'source': source_col,
            'targets': target_cols if isinstance(target_cols, list) else [target_cols]
        })
    
    def get_mappings(self):
        """获取所有映射关系"""
        return self.mappings
    
    def clear_mappings(self):
        """清除所有映射"""
        self.mappings = []
    
    def save_config(self, file_path):
        """保存映射配置到JSON文件
        
        Args:
            file_path (str): 配置文件路径
        """
        try:
            config = {
                'mappings': self.mappings
            }
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            raise Exception(f"保存配置失败: {str(e)}")
    
    def load_config(self, file_path):
        """从JSON文件加载映射配置
        
        Args:
            file_path (str): 配置文件路径
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self.mappings = config.get('mappings', [])
            return True
        except Exception as e:
            raise Exception(f"加载配置失败: {str(e)}")
    
    def to_dict(self):
        """将映射转换为字典格式"""
        return {'mappings': self.mappings}
    
    def from_dict(self, config_dict):
        """从字典格式恢复映射"""
        self.mappings = config_dict.get('mappings', [])


class ExcelDataProcessor:
    """Excel数据处理核心类"""
    
    def __init__(self):
        self.source_df = None
        self.template_df = None
        self.mapper = ExcelColumnMapper()
        
    def load_source_file(self, file_path, sheet_name):
        """加载发货明细表
        
        Args:
            file_path (str): 文件路径
            sheet_name (str): sheet名称
            
        Returns:
            bool: 加载成功返回True
        """
        try:
            self.source_df = pd.read_excel(file_path, sheet_name=sheet_name)
            return True
        except Exception as e:
            raise Exception(f"加载发货明细表失败: {str(e)}")
    
    def load_template_file(self, file_path, sheet_name):
        """加载模板表
        
        Args:
            file_path (str): 文件路径
            sheet_name (str): sheet名称
            
        Returns:
            bool: 加载成功返回True
        """
        try:
            self.template_df = pd.read_excel(file_path, sheet_name=sheet_name)
            return True
        except Exception as e:
            raise Exception(f"加载模板表失败: {str(e)}")
    
    def process_data(self, source_match_col, template_match_col):
        """处理数据并生成结果
        
        Args:
            source_match_col (str): 发货明细表的匹配列名
            template_match_col (str): 模板表的匹配列名
            
        Returns:
            pd.DataFrame: 处理后的数据
        """
        if self.source_df is None or self.template_df is None:
            raise Exception("请先加载源文件和模板文件")
        
        # 创建结果DataFrame，保持模板表的格式
        result_df = self.template_df.copy()
        
        # 用于存储所有新行的列表
        all_rows = []
        
        # 处理模板表中的每一行
        for _, template_row in self.template_df.iterrows():
            template_key = template_row[template_match_col]
            
            # 查找所有匹配的发货明细数据
            matched_data = self.source_df[self.source_df[source_match_col] == template_key]
            
            if not matched_data.empty:
                # 一对多关系：为每条匹配数据创建新行
                for _, source_row in matched_data.iterrows():
                    new_row = template_row.copy()
                    self._apply_mappings(new_row, source_row)
                    all_rows.append(new_row)
            # 如果模板表中的数据在发货明细表中没有匹配，直接跳过这行数据
            # 不再将模板行添加到结果中
        
        # 处理发货明细表中存在但模板表中没有的数据
        unmatched_source = self.source_df[~self.source_df[source_match_col].isin(
            self.template_df[template_match_col])]
        
        for _, source_row in unmatched_source.iterrows():
            new_row = pd.Series(index=self.template_df.columns, dtype='object')
            new_row[template_match_col] = source_row[source_match_col]
            self._apply_mappings(new_row, source_row)
            all_rows.append(new_row)
        
        # 创建最终DataFrame
        if all_rows:
            result_df = pd.DataFrame(all_rows)
            result_df = result_df.reset_index(drop=True)
        
        return result_df
    
    def _apply_mappings(self, target_row, source_row):
        """应用列映射到目标行
        
        Args:
            target_row (pd.Series): 目标行
            source_row (pd.Series): 源数据行
        """
        for mapping in self.mapper.get_mappings():
            source_col = mapping['source']
            target_cols = mapping['targets']
            
            if source_col in source_row.index:
                value = source_row[source_col]
                for target_col in target_cols:
                    if target_col in target_row.index:
                        target_row[target_col] = value
    
    def get_source_columns(self):
        """获取发货明细表的列名"""
        if self.source_df is not None:
            return list(self.source_df.columns)
        return []
    
    def get_template_columns(self):
        """获取模板表的列名"""
        if self.template_df is not None:
            return list(self.template_df.columns)
        return []


class ExcelFillerGUI:
    """图形用户界面类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel表格填写工具V1.2 by gengu312")
        self.root.geometry("700x950")
        self.root.minsize(700, 850)
        
        # 窗口居中显示
        self.root.update_idletasks()  # 更新窗口信息
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # 设置配置文件目录路径
        if getattr(sys, 'frozen', False):
            # 如果是打包后的exe文件
            self.config_dir = Path(sys.executable).parent / "configs"
        else:
            # 如果是开发环境
            self.config_dir = Path(__file__).parent / "configs"
        
        # 确保配置目录存在
        self.config_dir.mkdir(exist_ok=True)
        
        # 初始化配置文件列表
        self.config_files = []
        
        # 设置主题样式 - 现代简约风格
        self.style = ttk.Style()
        self.style.theme_use('clam')
        # 纯白色主题配色方案配清晰边框
        bg_primary = '#ffffff'      # 主背景白色
        bg_secondary = '#fafafa'    # 次要背景微灰白
        text_primary = '#000000'    # 主文本黑色
        text_secondary = '#333333'  # 次要文本深灰
        border_color = '#d0d0d0'    # 边框清晰灰色
        button_bg = '#ffffff'       # 按钮背景白色
        button_fg = '#000000'       # 按钮文字黑色
        
        # 自定义纯白色样式
        self.style.configure('Main.TFrame', background=bg_primary)
        self.style.configure('TLabelframe', background=bg_primary, relief='flat', borderwidth=1)
        self.style.configure('TLabelframe.Label', font=('Segoe UI', 11, 'bold'), foreground=text_primary, background=bg_primary)
        self.style.configure('TButton', font=('Segoe UI', 10), padding=10, relief='flat', background=button_bg, foreground=button_fg, 
                           borderwidth=1, highlightthickness=1, highlightbackground=border_color)
        self.style.map('TButton', background=[('active', '#f5f5f5'), ('disabled', '#f0f0f0')],
                      foreground=[('active', '#000000'), ('disabled', '#999999')])
        self.style.configure('TLabel', font=('Segoe UI', 10), background=bg_primary, foreground=text_primary)
        self.style.configure('TCombobox', font=('Segoe UI', 10), fieldbackground=bg_primary, background=bg_primary, foreground=text_primary)
        self.style.configure('White.TCombobox', fieldbackground=bg_primary, background=bg_primary, foreground=text_primary,
                          selectbackground='#f0f0f0', selectforeground='#000000')
        self.style.configure('Header.TLabel', font=('Segoe UI', 16, 'bold'), foreground=text_primary, background=bg_primary)
        
        self.processor = ExcelDataProcessor()
        self.source_file = None
        self.template_file = None
        self.source_sheet = None
        self.template_sheet = None
        self.output_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        # 设置根窗口背景为纯净白色
        self.root.configure(bg='#ffffff')
        
        # 创建主框架，使用更现代的间距
        main_frame = ttk.Frame(self.root, padding="25 20 25 20", style='Main.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重使框架可扩展
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # 创建内容框架
        content_frame = ttk.Frame(main_frame, style='Main.TFrame')
        content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content_frame.columnconfigure(0, weight=1)
        
        # 文件选择区域
        self.create_file_selection_frame(content_frame)
        
        # Sheet选择区域
        self.create_sheet_selection_frame(content_frame)
        
        # 匹配列配置区域
        self.create_match_column_frame(content_frame)
        
        # 列映射区域
        self.create_mapping_frame(content_frame)
        
        # 操作按钮区域
        self.create_action_frame(content_frame)
        
        # 进度显示区域
        self.create_progress_frame(content_frame)
    
    def create_file_selection_frame(self, parent):
        """创建现代文件选择框架"""
        frame = ttk.LabelFrame(parent, text="文件选择", padding="15")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        frame.configure(borderwidth=1, relief="solid")
        frame.columnconfigure(1, weight=1)
        
        # 发货明细表选择
        ttk.Label(frame, text="发货明细表").grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        file_frame1 = ttk.Frame(frame)
        file_frame1.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 15))
        file_frame1.columnconfigure(0, weight=1)
        
        self.source_file_label = ttk.Label(file_frame1, text="未选择文件", width=50, 
                                         background='#ffffff', padding=8, relief='solid', 
                                         borderwidth=1, foreground='#666666')
        self.source_file_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        tk.Button(frame, text="选择文件", command=self.select_source_file,
                  background='#ffffff', foreground='#000000',
                  activebackground='#f0f0f0', activeforeground='#000000',
                  relief='solid', borderwidth=1, padx=10, pady=5).grid(row=0, column=2)
        
        # 模板表选择
        ttk.Label(frame, text="模板表").grid(row=1, column=0, sticky=tk.W, padx=(0, 15), pady=(10, 0))
        file_frame2 = ttk.Frame(frame)
        file_frame2.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 15), pady=(10, 0))
        file_frame2.columnconfigure(0, weight=1)
        
        self.template_file_label = ttk.Label(file_frame2, text="未选择文件", width=50,
                                           background='#ffffff', padding=8, relief='solid', 
                                           borderwidth=1, foreground='#666666')
        self.template_file_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        tk.Button(frame, text="选择文件", command=self.select_template_file,
                  background='#ffffff', foreground='#000000',
                  activebackground='#f0f0f0', activeforeground='#000000',
                  relief='solid', borderwidth=1, padx=10, pady=5).grid(row=1, column=2, pady=(10, 0))
    
    def create_sheet_selection_frame(self, parent):
        """创建现代sheet选择框架"""
        frame = ttk.LabelFrame(parent, text="工作表选择", padding="15")
        frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        frame.configure(borderwidth=1, relief="solid")
        
        # 创建两列布局
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        
        # 发货明细表sheet
        ttk.Label(frame, text="发货明细表").grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        self.source_sheet_combo = ttk.Combobox(frame, state="readonly", width=30,
                                             style='White.TCombobox')
        self.source_sheet_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        self.source_sheet_combo.bind("<<ComboboxSelected>>", self.on_source_sheet_selected)
        
        # 模板表sheet
        ttk.Label(frame, text="模板表").grid(row=0, column=1, sticky=tk.W)
        self.template_sheet_combo = ttk.Combobox(frame, state="readonly", width=30,
                                               style='White.TCombobox')
        self.template_sheet_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(5, 0))
        self.template_sheet_combo.bind("<<ComboboxSelected>>", self.on_template_sheet_selected)
    
    def create_match_column_frame(self, parent):
        """创建匹配列配置框架"""
        frame = ttk.LabelFrame(parent, text="匹配列配置", padding="10")
        frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        frame.configure(borderwidth=1, relief="solid")
        
        # 创建两列布局
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        
        ttk.Label(frame, text="发货明细表匹配列:").grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        self.source_match_combo = ttk.Combobox(frame, state="readonly", width=35, 
                                             style='White.TCombobox')
        self.source_match_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Label(frame, text="模板表匹配列:").grid(row=0, column=1, sticky=tk.W)
        self.template_match_combo = ttk.Combobox(frame, state="readonly", width=35,
                                               style='White.TCombobox')
        self.template_match_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def create_mapping_frame(self, parent):
        """创建列映射框架"""
        frame = ttk.LabelFrame(parent, text="列映射配置", padding="10")
        frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        frame.configure(borderwidth=1, relief="solid")
        frame.columnconfigure(0, weight=1)
        
        # 快速配置选择区域
        quick_config_frame = ttk.Frame(frame)
        quick_config_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        quick_config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(quick_config_frame, text="快速选择配置:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # 配置文件下拉框
        self.quick_config_combo = ttk.Combobox(quick_config_frame, state="readonly", width=35,
                                             style='White.TCombobox')
        self.quick_config_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.quick_config_combo.bind("<<ComboboxSelected>>", self.on_quick_config_selected)
        
        # 刷新按钮
        refresh_btn = tk.Button(quick_config_frame, text="刷新", command=self.refresh_config_files,
                              background='#ffffff', foreground='#000000',
                              activebackground='#f0f0f0', activeforeground='#000000',
                              relief='solid', borderwidth=1, padx=10, pady=5)
        refresh_btn.grid(row=0, column=2)
        
        # 加载配置文件列表
        self.refresh_config_files()
        
        # 映射配置区域 - 原来的代码，行号从1改为2
        mapping_frame = ttk.Frame(frame)
        mapping_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 设置列权重使组件可扩展
        for i in range(4):
            mapping_frame.columnconfigure(i, weight=1)
        
        # 发货明细表列
        ttk.Label(mapping_frame, text="发货明细表列").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.mapping_source_combo = ttk.Combobox(mapping_frame, state="readonly", width=20,
                                               style='White.TCombobox')
        self.mapping_source_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        # 模板表列1
        ttk.Label(mapping_frame, text="模板表列1 (必填)").grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        self.mapping_target_combo1 = ttk.Combobox(mapping_frame, state="readonly", width=20,
                                                 style='White.TCombobox')
        self.mapping_target_combo1.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        # 模板表列2（可选）
        ttk.Label(mapping_frame, text="模板表列2 (可选)").grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        self.mapping_target_combo2 = ttk.Combobox(mapping_frame, state="readonly", width=20,
                                                style='White.TCombobox')
        self.mapping_target_combo2.grid(row=1, column=2, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        # 添加映射按钮
        tk.Button(mapping_frame, text="添加映射", command=self.add_mapping,
                   background='#ffffff', foreground='#000000',
                   activebackground='#f0f0f0', activeforeground='#000000',
                   relief='solid', borderwidth=1, padx=10, pady=5).grid(row=1, column=3, sticky=(tk.W, tk.E), padx=(5, 0))
        
        # 映射列表 - 使用更美观的列表框
        list_frame = ttk.Frame(frame)
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        
        # 创建带滚动条的列表框
        self.mapping_listbox = tk.Listbox(list_frame, height=6, width=90, font=('Consolas', 10),
                                        selectbackground='#0078d4', selectforeground='white',
                                        background='#fafafa', relief='groove', borderwidth=1)
        self.mapping_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.mapping_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.mapping_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 映射操作按钮区域 - 带边框的按钮布局
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        buttons = [
            ("删除选中", self.remove_mapping),
            ("保存配置", self.save_mapping_config),
            ("加载配置", self.load_mapping_config),
            ("清空映射", self.clear_mappings)
        ]
        
        for i, (text, command) in enumerate(buttons):
            btn = tk.Button(button_frame, text=text, command=command,
                          background='#ffffff', foreground='#000000',
                          activebackground='#f0f0f0', activeforeground='#000000',
                          relief='solid', borderwidth=1, padx=10, pady=5)
            btn.grid(row=0, column=i, padx=(0 if i == 0 else 5, 5), sticky=(tk.W, tk.E))
            button_frame.columnconfigure(i, weight=1)
    
    def create_action_frame(self, parent):
        """创建操作按钮框架"""
        frame = ttk.LabelFrame(parent, text="操作控制", padding="10")
        frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        frame.configure(borderwidth=1, relief="solid")
        frame.columnconfigure(0, weight=1)
        
        # 创建按钮容器
        button_container = ttk.Frame(frame)
        button_container.grid(row=0, column=0)
        
        # 主要操作按钮 - 带边框的白色按钮
        main_buttons = [
            ("开始处理", self.start_processing),
            ("清除配置", self.clear_all)
        ]
        
        for i, (text, command) in enumerate(main_buttons):
            btn = tk.Button(button_container, text=text, command=command,
                          background='#ffffff', foreground='#000000',
                          activebackground='#f0f0f0', activeforeground='#000000',
                          relief='solid', borderwidth=1, padx=15, pady=8)
            btn.grid(row=0, column=i, padx=(0 if i == 0 else 10, 0))
    
    def create_progress_frame(self, parent):
        """创建进度显示框架"""
        frame = ttk.LabelFrame(parent, text="处理进度", padding="10")
        frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        frame.configure(borderwidth=1, relief="solid")
        frame.columnconfigure(0, weight=1)
        
        # 进度条区域
        progress_frame = ttk.Frame(frame)
        progress_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        progress_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100,
                                          length=400, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(5, 5))
        
        # 状态标签
        self.status_label = ttk.Label(progress_frame, text="✅ 就绪 - 等待开始处理", 
                                    font=('Microsoft YaHei', 10, 'bold'), foreground='#28a745')
        self.status_label.grid(row=1, column=0, pady=(5, 0))
    
    def select_source_file(self):
        """选择发货明细表"""
        file_path = filedialog.askopenfilename(
            title="选择发货明细表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.source_file = file_path
            self.source_file_label.config(text=basename(file_path))
            self.load_sheets(file_path, "source")
    
    def select_template_file(self):
        """选择模板表"""
        file_path = filedialog.askopenfilename(
            title="选择模板表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.template_file = file_path
            self.template_file_label.config(text=basename(file_path))
            self.load_sheets(file_path, "template")
    
    def load_sheets(self, file_path, file_type):
        """加载Excel文件的sheet列表，并自动选择第一个sheet"""
        try:
            excel_file = pd.ExcelFile(file_path)
            sheets = excel_file.sheet_names
            
            if file_type == "source":
                self.source_sheet_combo['values'] = sheets
                if sheets:
                    self.source_sheet_combo.current(0)
                    # 自动触发sheet选择事件
                    self.source_sheet = sheets[0]
                    if self.source_file:
                        self.processor.load_source_file(self.source_file, self.source_sheet)
                        columns = self.processor.get_source_columns()
                        self.source_match_combo['values'] = columns
                        self.mapping_source_combo['values'] = columns
            else:
                self.template_sheet_combo['values'] = sheets
                if sheets:
                    self.template_sheet_combo.current(0)
                    # 自动触发sheet选择事件
                    self.template_sheet = sheets[0]
                    if self.template_file:
                        self.processor.load_template_file(self.template_file, self.template_sheet)
                        columns = self.processor.get_template_columns()
                        self.template_match_combo['values'] = columns
                        self.mapping_target_combo1['values'] = columns
                        self.mapping_target_combo2['values'] = columns
        except Exception as e:
            messagebox.showerror("错误", f"加载sheet列表失败: {str(e)}")
    
    def on_source_sheet_selected(self, event):
        """发货明细表sheet选择事件"""
        self.source_sheet = self.source_sheet_combo.get()
        if self.source_file and self.source_sheet:
            try:
                self.processor.load_source_file(self.source_file, self.source_sheet)
                columns = self.processor.get_source_columns()
                self.source_match_combo['values'] = columns
                self.mapping_source_combo['values'] = columns
            except Exception as e:
                messagebox.showerror("错误", str(e))
    
    def on_template_sheet_selected(self, event):
        """模板表sheet选择事件"""
        self.template_sheet = self.template_sheet_combo.get()
        if self.template_file and self.template_sheet:
            try:
                self.processor.load_template_file(self.template_file, self.template_sheet)
                columns = self.processor.get_template_columns()
                self.template_match_combo['values'] = columns
                self.mapping_target_combo1['values'] = columns
                self.mapping_target_combo2['values'] = columns
            except Exception as e:
                messagebox.showerror("错误", str(e))
    
    def add_mapping(self):
        """添加列映射，支持一对二映射"""
        source_col = self.mapping_source_combo.get()
        target_col1 = self.mapping_target_combo1.get()
        target_col2 = self.mapping_target_combo2.get()
        
        if not source_col or not target_col1:
            messagebox.showwarning("警告", "请选择源列和模板表列1")
            return
            
        # 添加第一个映射
        mapping_text1 = f"{source_col} -> {target_col1}"
        self.mapping_listbox.insert(tk.END, mapping_text1)
        self.processor.mapper.add_mapping(source_col, target_col1)
        
        # 如果选择了模板表列2，添加第二个映射
        if target_col2:
            mapping_text2 = f"{source_col} -> {target_col2}"
            self.mapping_listbox.insert(tk.END, mapping_text2)
            self.processor.mapper.add_mapping(source_col, target_col2)
            
        # 清空选择
        self.mapping_source_combo.set('')
        self.mapping_target_combo1.set('')
        self.mapping_target_combo2.set('')
    
    def remove_mapping(self):
        """删除选中的映射"""
        selection = self.mapping_listbox.curselection()
        if selection:
            index = selection[0]
            self.mapping_listbox.delete(index)
            mappings = self.processor.mapper.get_mappings()
            if index < len(mappings):
                del mappings[index]
    
    def start_processing(self):
        """开始处理数据"""
        if not self.validate_inputs():
            return
        
        # 强制选择导出路径
        if not self.output_path:
            self.output_path = filedialog.askdirectory(title="选择导出文件夹")
            if not self.output_path:  # 用户取消了选择
                return
        
        # 创建线程处理数据，避免界面卡顿
        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()
    
    def validate_inputs(self):
        """验证输入是否完整"""
        if not self.source_file or not self.template_file:
            messagebox.showwarning("警告", "请选择发货明细表和模板表")
            return False
        
        if not self.source_sheet or not self.template_sheet:
            messagebox.showwarning("警告", "请选择sheet")
            return False
        
        if not self.source_match_combo.get() or not self.template_match_combo.get():
            messagebox.showwarning("警告", "请选择匹配列")
            return False
        
        return True
    
    def process_data(self):
        """处理数据"""
        try:
            self.update_status("正在处理数据...", 10)
            
            source_match_col = self.source_match_combo.get()
            template_match_col = self.template_match_combo.get()
            
            result_df = self.processor.process_data(source_match_col, template_match_col)
            
            self.update_status("正在保存结果...", 80)
            
            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"填写表_{timestamp}.xlsx"
            
            # 确定输出路径
            if self.output_path:
                output_file = join(self.output_path, output_filename)
            else:
                output_file = output_filename
            
            # 保存结果
            result_df.to_excel(output_file, index=False)
            
            self.update_status("处理完成！", 100)
            
            # 获取导出文件夹路径
            output_dir = dirname(output_file)
            
            # 询问用户是否打开文件夹
            if messagebox.askyesno("处理完成", f"文件已成功保存！\n\n文件名: {basename(output_file)}\n\n位置: {output_dir}\n\n是否打开所在文件夹？"):
                try:
                    # 使用Windows资源管理器打开文件夹并选中文件
                    # 处理路径中的斜杠问题，确保使用反斜杠
                    normalized_path = normpath(output_file)
                    subprocess.run(['explorer', '/select,', normalized_path], shell=True, check=True)
                except Exception as e:
                    # 如果无法选中文件，尝试打开文件夹
                    try:
                        normalized_dir = normpath(output_dir)
                        subprocess.run(['explorer', normalized_dir], shell=True, check=True)
                    except Exception as e2:
                        # 最后尝试使用start命令
                        try:
                            startfile(output_dir)
                        except Exception as e3:
                            messagebox.showwarning("提示", f"无法自动打开文件夹: {str(e3)}\n\n请手动打开文件夹: {output_dir}")
            
        except Exception as e:
            self.update_status("处理失败", 0)
            messagebox.showerror("错误", f"处理失败: {str(e)}")
    
    def update_status(self, message, progress):
        """更新状态和进度"""
        self.status_label.config(text=message)
        self.progress_var.set(progress)
        self.root.update_idletasks()
    

    
    def clear_all(self):
        """清除所有配置"""
        self.source_file = None
        self.template_file = None
        self.source_sheet = None
        self.template_sheet = None
        self.output_path = None
        
        self.source_file_label.config(text="未选择")
        self.template_file_label.config(text="未选择")
        self.output_path_label.config(text="未选择（默认：程序目录）")
        self.source_sheet_combo.set('')
        self.template_sheet_combo.set('')
        self.source_match_combo.set('')
        self.template_match_combo.set('')
        self.mapping_source_combo.set('')
        self.mapping_target_combo1.set('')
        self.mapping_target_combo2.set('')
        self.mapping_listbox.delete(0, tk.END)
        
        self.processor.mapper.clear_mappings()
        self.update_status("就绪", 0)
    
    def refresh_config_files(self):
        """刷新配置文件列表"""
        try:
            # 获取所有JSON配置文件
            self.config_files = [f.name for f in self.config_dir.glob("*.json")]
            self.quick_config_combo['values'] = self.config_files
            if self.config_files:
                self.quick_config_combo.set("选择配置文件...")
            else:
                self.quick_config_combo.set("无配置文件")
        except Exception as e:
            messagebox.showerror("错误", f"刷新配置文件失败: {str(e)}")
    
    def on_quick_config_selected(self, event):
        """快速配置选择事件"""
        selected_file = self.quick_config_combo.get()
        if selected_file and selected_file != "选择配置文件..." and selected_file != "无配置文件":
            config_path = self.config_dir / selected_file
            self.load_mapping_config_from_path(str(config_path))
    
    def load_mapping_config_from_path(self, file_path):
        """从指定路径加载映射配置"""
        try:
            self.processor.mapper.load_config(file_path)
            self.update_mapping_list()
            messagebox.showinfo("成功", f"配置已成功加载: {Path(file_path).name}")
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败: {str(e)}")
    
    def save_mapping_config(self):
        """保存映射配置到配置目录"""
        if not self.processor.mapper.get_mappings():
            messagebox.showwarning("警告", "没有可保存的映射配置")
            return
        
        # 弹出输入框获取文件名
        filename = tk.simpledialog.askstring("保存配置", "请输入配置文件名称:", 
                                           initialvalue="mapping_config.json")
        if not filename:
            return
        
        # 确保文件名以.json结尾
        if not filename.endswith('.json'):
            filename += '.json'
        
        file_path = self.config_dir / filename
        
        try:
            self.processor.mapper.save_config(str(file_path))
            self.refresh_config_files()  # 刷新配置文件列表
            messagebox.showinfo("成功", f"配置已保存到: {filename}")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def load_mapping_config(self):
        """从配置目录加载映射配置"""
        file_path = filedialog.askopenfilename(
            title="选择映射配置文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            initialdir=str(self.config_dir)
        )
        
        if file_path:
            self.load_mapping_config_from_path(file_path)
    
    def update_mapping_list(self):
        """更新映射列表显示"""
        self.mapping_listbox.delete(0, tk.END)
        for mapping in self.processor.mapper.get_mappings():
            source = mapping['source']
            targets = ", ".join(mapping['targets'])
            self.mapping_listbox.insert(tk.END, f"{source} → {targets}")
    
    def clear_mappings(self):
        """仅清空映射配置"""
        self.processor.mapper.clear_mappings()
        self.update_mapping_list()
        messagebox.showinfo("成功", "映射配置已清空")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ExcelFillerGUI()
    app.run()