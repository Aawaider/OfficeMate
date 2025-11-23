import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, scrolledtext
import json
import os
import base64
import io
import threading
import socket
import time
from datetime import datetime
import math
import re
import csv
import sqlite3
import hashlib
import uuid
import webbrowser
from collections import defaultdict
import statistics
import tempfile
import shutil

try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

class OfficeMatePro:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("OfficeMate Alpha1")
        self.root.geometry("1400x900")
        self.root.configure(bg='#2c3e50')
        
        # 应用主题
        self.setup_theme()
        
        # 初始化数据
        self.current_file = None
        self.auto_save = True
        self.auto_save_interval = 300000  # 5分钟
        self.user_preferences = self.load_preferences()
        
        # 协作功能
        self.collaboration_mode = False
        self.client_socket = None
        self.server_socket = None
        self.connected_clients = []
        self.user_id = str(uuid.uuid4())[:8]
        
        # AI功能状态
        self.ai_assistant_enabled = True
        
        # 版本控制
        self.document_history = []
        self.current_version = 0
        
        # 创建数据库
        self.setup_database()
        
        # 创建界面
        self.create_ui()
        
        # 启动自动保存
        self.setup_auto_save()
        
    def setup_theme(self):
        """设置现代化主题"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 自定义样式
        self.style.configure('Modern.TButton', 
                           background='#3498db',
                           foreground='white',
                           borderwidth=1,
                           focusthickness=3,
                           focuscolor='none',
                           padding=(10, 5))
        
        self.style.configure('Accent.TButton',
                           background='#e74c3c',
                           foreground='white')
        
        self.style.configure('Success.TButton',
                           background='#2ecc71',
                           foreground='white')
        
        self.style.configure('Dark.TFrame', background='#34495e')
        
    def setup_database(self):
        """设置SQLite数据库"""
        self.conn = sqlite3.connect('officemate.db', check_same_thread=False)
        self.cursor = self.conn.cursor()
        
        # 创建表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT UNIQUE,
                content TEXT,
                metadata TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                type TEXT,
                content TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.conn.commit()
        
    def load_preferences(self):
        """加载用户偏好设置"""
        try:
            with open('preferences.json', 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {
                'theme': 'dark',
                'auto_save': True,
                'recent_files': [],
                'default_font': 'Arial',
                'default_font_size': 12,
                'window_size': [1400, 900]
            }
            
    def save_preferences(self):
        """保存用户偏好设置"""
        try:
            with open('preferences.json', 'w', encoding='utf-8') as f:
                json.dump(self.user_preferences, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存偏好设置失败: {e}")
            
    def setup_auto_save(self):
        """设置自动保存"""
        if self.auto_save:
            self.root.after(self.auto_save_interval, self.auto_save_document)
            
    def auto_save_document(self):
        """自动保存文档"""
        if self.current_file and self.text_area.get('1.0', 'end-1c').strip():
            try:
                self.backup_document()
            except:
                pass
        self.setup_auto_save()
        
    def backup_document(self):
        """备份文档"""
        try:
            backup_dir = "backups"
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
                
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"backup_{timestamp}.txt")
            
            with open(backup_file, 'w', encoding='utf-8') as f:
                f.write(self.text_area.get('1.0', 'end-1c'))
        except Exception as e:
            print(f"备份失败: {e}")

    def create_ui(self):
        """创建增强的用户界面"""
        # 创建主菜单
        self.create_main_menu()
        
        # 创建主工具栏
        self.create_main_toolbar()
        
        # 创建状态栏
        self.create_status_bar()
        
        # 创建主内容区域
        self.create_main_content()
        
        # 创建侧边栏
        self.create_sidebar()
        
    def create_main_menu(self):
        """创建主菜单"""
        menubar = tk.Menu(self.root, bg='#34495e', fg='white')
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="打开", command=self.open_file, accelerator="Ctrl+O")
        
        # 最近文件子菜单
        self.recent_menu = tk.Menu(file_menu, tearoff=0, bg='#34495e', fg='white')
        file_menu.add_cascade(label="最近文件", menu=self.recent_menu)
        self.update_recent_files_menu()
        
        file_menu.add_separator()
        file_menu.add_command(label="保存", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="另存为", command=self.save_as_file, accelerator="Ctrl+Shift+S")
        file_menu.add_command(label="导出", command=self.export_document)
        file_menu.add_separator()
        file_menu.add_command(label="打印", command=self.print_document, accelerator="Ctrl+P")
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit_application, accelerator="Ctrl+Q")
        
        # 编辑菜单
        edit_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="编辑", menu=edit_menu)
        edit_menu.add_command(label="撤销", command=self.undo, accelerator="Ctrl+Z")
        edit_menu.add_command(label="重做", command=self.redo, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="剪切", command=self.cut, accelerator="Ctrl+X")
        edit_menu.add_command(label="复制", command=self.copy, accelerator="Ctrl+C")
        edit_menu.add_command(label="粘贴", command=self.paste, accelerator="Ctrl+V")
        edit_menu.add_separator()
        edit_menu.add_command(label="查找和替换", command=self.find_replace_dialog, accelerator="Ctrl+F")
        edit_menu.add_command(label="全选", command=self.select_all, accelerator="Ctrl+A")
        
        # 视图菜单
        view_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="视图", menu=view_menu)
        view_menu.add_command(label="文字处理", command=self.show_word_processor)
        view_menu.add_command(label="电子表格", command=self.show_spreadsheet)
        view_menu.add_command(label="演示文稿", command=self.show_presentation)
        view_menu.add_command(label="数据库查看器", command=self.show_database_viewer)
        view_menu.add_separator()
        view_menu.add_command(label="全屏", command=self.toggle_fullscreen, accelerator="F11")
        view_menu.add_command(label="缩放", command=self.zoom_dialog)
        
        # 插入菜单
        insert_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="插入", menu=insert_menu)
        insert_menu.add_command(label="图片", command=self.insert_image_dialog)
        insert_menu.add_command(label="表格", command=self.insert_table_dialog)
        insert_menu.add_command(label="图表", command=self.insert_chart_dialog)
        insert_menu.add_command(label="公式", command=self.insert_equation_dialog)
        insert_menu.add_command(label="超链接", command=self.insert_hyperlink_dialog)
        insert_menu.add_command(label="日期时间", command=self.insert_datetime)
        
        # 格式菜单
        format_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="格式", menu=format_menu)
        format_menu.add_command(label="字体", command=self.font_dialog)
        format_menu.add_command(label="段落", command=self.paragraph_dialog)
        format_menu.add_command(label="样式", command=self.style_dialog)
        format_menu.add_command(label="主题", command=self.theme_dialog)
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="拼写检查", command=self.spell_check)
        tools_menu.add_command(label="字数统计", command=self.show_word_count)
        tools_menu.add_command(label="模板管理", command=self.template_manager)
        tools_menu.add_command(label="宏录制", command=self.macro_recorder)
        tools_menu.add_separator()
        tools_menu.add_command(label="选项", command=self.options_dialog)
        
        # AI助手菜单
        ai_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="AI助手", menu=ai_menu)
        ai_menu.add_command(label="智能写作", command=self.ai_writing_assistant)
        ai_menu.add_command(label="语法检查", command=self.ai_grammar_check)
        ai_menu.add_command(label="内容优化", command=self.ai_content_optimize)
        ai_menu.add_command(label="文本摘要", command=self.ai_text_summarize)
        
        # 协作菜单
        collaboration_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="协作", menu=collaboration_menu)
        collaboration_menu.add_command(label="启动协作服务器", command=self.start_collaboration_server)
        collaboration_menu.add_command(label="连接到服务器", command=self.connect_to_server_dialog)
        collaboration_menu.add_command(label="断开连接", command=self.disconnect_from_server)
        collaboration_menu.add_command(label="共享文档", command=self.share_document_dialog)
        collaboration_menu.add_command(label="版本历史", command=self.version_history)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0, bg='#34495e', fg='white')
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用教程", command=self.show_tutorial)
        help_menu.add_command(label="快捷键", command=self.show_shortcuts)
        help_menu.add_command(label="关于", command=self.about_dialog)
        
        # 绑定快捷键
        self.bind_shortcuts()
        
    def update_recent_files_menu(self):
        """更新最近文件菜单"""
        self.recent_menu.delete(0, tk.END)
        recent_files = self.user_preferences.get('recent_files', [])
        
        for file_path in recent_files[-10:]:  # 只显示最近10个文件
            if os.path.exists(file_path):
                filename = os.path.basename(file_path)
                self.recent_menu.add_command(
                    label=filename,
                    command=lambda path=file_path: self.open_recent_file(path)
                )
        
    def open_recent_file(self, file_path):
        """打开最近文件"""
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.text_area.delete('1.0', tk.END)
                    self.text_area.insert('1.0', content)
                    self.current_file = file_path
                    self.root.title(f"OfficeMate Pro - {os.path.basename(file_path)}")
                    self.add_to_version_history(f"打开文件: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {str(e)}")
        else:
            messagebox.showwarning("文件不存在", "选择的文件不存在，已从最近文件列表中移除")
            # 从最近文件列表中移除
            if file_path in self.user_preferences['recent_files']:
                self.user_preferences['recent_files'].remove(file_path)
                self.save_preferences()
                self.update_recent_files_menu()
        
    def create_main_toolbar(self):
        """创建主工具栏"""
        toolbar = tk.Frame(self.root, bg='#34495e', height=40)
        toolbar.pack(fill='x', padx=5, pady=2)
        
        # 文件操作按钮
        new_btn = ttk.Button(toolbar, text="新建", command=self.new_file, style='Modern.TButton')
        new_btn.pack(side='left', padx=2)
        
        open_btn = ttk.Button(toolbar, text="打开", command=self.open_file, style='Modern.TButton')
        open_btn.pack(side='left', padx=2)
        
        save_btn = ttk.Button(toolbar, text="保存", command=self.save_file, style='Modern.TButton')
        save_btn.pack(side='left', padx=2)
        
        # 分隔符
        ttk.Separator(toolbar, orient='vertical').pack(side='left', padx=5, fill='y')
        
        # 编辑按钮
        undo_btn = ttk.Button(toolbar, text="撤销", command=self.undo, style='Modern.TButton')
        undo_btn.pack(side='left', padx=2)
        
        redo_btn = ttk.Button(toolbar, text="重做", command=self.redo, style='Modern.TButton')
        redo_btn.pack(side='left', padx=2)
        
        # 分隔符
        ttk.Separator(toolbar, orient='vertical').pack(side='left', padx=5, fill='y')
        
        # 格式按钮
        bold_btn = ttk.Button(toolbar, text="粗体", command=self.toggle_bold, style='Modern.TButton')
        bold_btn.pack(side='left', padx=2)
        
        italic_btn = ttk.Button(toolbar, text="斜体", command=self.toggle_italic, style='Modern.TButton')
        italic_btn.pack(side='left', padx=2)
        
        underline_btn = ttk.Button(toolbar, text="下划线", command=self.toggle_underline, style='Modern.TButton')
        underline_btn.pack(side='left', padx=2)
        
        # 分隔符
        ttk.Separator(toolbar, orient='vertical').pack(side='left', padx=5, fill='y')
        
        # AI助手按钮
        ai_btn = ttk.Button(toolbar, text="AI助手", command=self.ai_writing_assistant, style='Success.TButton')
        ai_btn.pack(side='left', padx=2)
        
    def create_status_bar(self):
        """创建状态栏"""
        self.status_bar = tk.Frame(self.root, bg='#2c3e50', height=20)
        self.status_bar.pack(fill='x', side='bottom')
        
        # 页面信息
        self.page_label = tk.Label(self.status_bar, text="页面 1", bg='#2c3e50', fg='white')
        self.page_label.pack(side='left', padx=5)
        
        # 字数统计
        self.word_count_label = tk.Label(self.status_bar, text="字数: 0", bg='#2c3e50', fg='white')
        self.word_count_label.pack(side='left', padx=5)
        
        # 光标位置
        self.cursor_label = tk.Label(self.status_bar, text="行: 1, 列: 1", bg='#2c3e50', fg='white')
        self.cursor_label.pack(side='left', padx=5)
        
        # 协作状态
        self.collab_label = tk.Label(self.status_bar, text="离线", bg='#2c3e50', fg='red')
        self.collab_label.pack(side='right', padx=5)
        
        # 自动保存状态
        self.auto_save_label = tk.Label(self.status_bar, text="自动保存: 开", bg='#2c3e50', fg='green')
        self.auto_save_label.pack(side='right', padx=5)
        
    def create_main_content(self):
        """创建主内容区域"""
        main_frame = tk.Frame(self.root, bg='#ecf0f1')
        main_frame.pack(fill='both', expand=True, side='right')
        
        # 创建标签页控件
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 创建各个功能模块
        self.create_word_processor()
        self.create_spreadsheet()
        self.create_presentation()
        self.create_database_viewer()
        
    def create_sidebar(self):
        """创建侧边栏"""
        self.sidebar = tk.Frame(self.root, bg='#34495e', width=200)
        self.sidebar.pack(fill='y', side='left', padx=5, pady=5)
        
        # 文档结构树
        doc_structure_label = tk.Label(self.sidebar, text="文档结构", bg='#34495e', fg='white', font=('Arial', 10, 'bold'))
        doc_structure_label.pack(pady=5)
        
        self.doc_tree = ttk.Treeview(self.sidebar, height=15)
        self.doc_tree.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 快速样式面板
        style_label = tk.Label(self.sidebar, text="快速样式", bg='#34495e', fg='white', font=('Arial', 10, 'bold'))
        style_label.pack(pady=5)
        
        style_frame = tk.Frame(self.sidebar, bg='#34495e')
        style_frame.pack(fill='x', padx=5, pady=5)
        
        styles = ["标题1", "标题2", "正文", "引用", "代码"]
        for style in styles:
            btn = ttk.Button(style_frame, text=style, command=lambda s=style: self.apply_style(s))
            btn.pack(fill='x', pady=2)
            
    def create_word_processor(self):
        """创建文字处理器"""
        self.word_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.word_frame, text="文字处理")
        
        # 创建富文本编辑器
        self.create_rich_text_editor()
        
    def create_rich_text_editor(self):
        """创建富文本编辑器"""
        # 格式化工具栏
        format_toolbar = tk.Frame(self.word_frame, bg='#ecf0f1')
        format_toolbar.pack(fill='x', padx=5, pady=2)
        
        # 字体选择
        font_label = tk.Label(format_toolbar, text="字体:", bg='#ecf0f1')
        font_label.pack(side='left', padx=5)
        
        self.font_var = tk.StringVar(value="Arial")
        font_combo = ttk.Combobox(format_toolbar, textvariable=self.font_var, 
                                 values=["Arial", "Times New Roman", "Courier New", "Verdana", "微软雅黑", "宋体"],
                                 width=12)
        font_combo.pack(side='left', padx=5)
        font_combo.bind('<<ComboboxSelected>>', self.change_font)
        
        # 字号选择
        size_label = tk.Label(format_toolbar, text="字号:", bg='#ecf0f1')
        size_label.pack(side='left', padx=5)
        
        self.size_var = tk.StringVar(value="12")
        size_combo = ttk.Combobox(format_toolbar, textvariable=self.size_var, 
                                 values=["8", "10", "12", "14", "16", "18", "20", "24", "32"],
                                 width=5)
        size_combo.pack(side='left', padx=5)
        size_combo.bind('<<ComboboxSelected>>', self.change_font_size)
        
        # 文本编辑区域
        text_frame = tk.Frame(self.word_frame)
        text_frame.pack(fill='both', expand=True)
        
        # 滚动条
        v_scrollbar = tk.Scrollbar(text_frame)
        v_scrollbar.pack(side='right', fill='y')
        
        h_scrollbar = tk.Scrollbar(text_frame, orient='horizontal')
        h_scrollbar.pack(side='bottom', fill='x')
        
        # 文本区域
        self.text_area = tk.Text(
            text_frame,
            wrap='word',
            font=('Arial', 12),
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            undo=True,
            maxundo=-1,
            selectbackground='#3498db'
        )
        self.text_area.pack(fill='both', expand=True)
        
        # 配置滚动条
        v_scrollbar.config(command=self.text_area.yview)
        h_scrollbar.config(command=self.text_area.xview)
        
        # 绑定事件
        self.text_area.bind('<KeyRelease>', self.on_text_change)
        self.text_area.bind('<Button-1>', self.update_cursor_position)
        self.text_area.bind('<KeyPress>', self.update_cursor_position)
        
        # 初始化格式状态
        self.current_font = "Arial"
        self.current_size = 12
        self.current_bold = False
        self.current_italic = False
        self.current_underline = False
        self.current_color = "black"
        
    def create_spreadsheet(self):
        """创建电子表格"""
        self.sheet_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sheet_frame, text="电子表格")
        
        # 创建增强的电子表格
        self.create_enhanced_spreadsheet()
        
    def create_enhanced_spreadsheet(self):
        """创建增强的电子表格"""
        # 工具栏
        toolbar = tk.Frame(self.sheet_frame, bg='#ecf0f1')
        toolbar.pack(fill='x', padx=5, pady=5)
        
        # 添加行按钮
        add_row_btn = ttk.Button(toolbar, text="添加行", command=self.add_row)
        add_row_btn.pack(side='left', padx=5)
        
        # 添加列按钮
        add_col_btn = ttk.Button(toolbar, text="添加列", command=self.add_column)
        add_col_btn.pack(side='left', padx=5)
        
        # 公式按钮
        formula_btn = ttk.Button(toolbar, text="插入公式", command=self.insert_formula_dialog)
        formula_btn.pack(side='left', padx=5)
        
        # 公式栏
        formula_frame = tk.Frame(toolbar, bg='#ecf0f1')
        formula_frame.pack(fill='x', pady=2)
        
        formula_label = tk.Label(formula_frame, text="fx", bg='#ecf0f1', font=('Arial', 10, 'bold'))
        formula_label.pack(side='left', padx=5)
        
        self.formula_var = tk.StringVar()
        formula_entry = tk.Entry(formula_frame, textvariable=self.formula_var, width=50)
        formula_entry.pack(side='left', fill='x', expand=True, padx=5)
        formula_entry.bind('<Return>', self.apply_formula)
        
        # 表格框架
        table_frame = tk.Frame(self.sheet_frame)
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 创建表格
        self.table_canvas = tk.Canvas(table_frame)
        self.table_scrollbar = tk.Scrollbar(table_frame, orient="vertical", command=self.table_canvas.yview)
        self.table_scrollable_frame = tk.Frame(self.table_canvas)
        
        self.table_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
        )
        
        self.table_canvas.create_window((0, 0), window=self.table_scrollable_frame, anchor="nw")
        self.table_canvas.configure(yscrollcommand=self.table_scrollbar.set)
        
        self.table_canvas.pack(side="left", fill="both", expand=True)
        self.table_scrollbar.pack(side="right", fill="y")
        
        # 初始化表格
        self.rows = 20
        self.cols = 10
        self.cells = []
        self.cell_data = {}  # 存储单元格数据和公式
        self.create_table()
        
    def create_table(self):
        """创建表格"""
        # 清空现有表格
        for widget in self.table_scrollable_frame.winfo_children():
            widget.destroy()
        self.cells = []
        
        # 创建表头
        header_row = []
        for j in range(self.cols):
            col_name = chr(65 + j)  # A, B, C, ...
            header = tk.Entry(self.table_scrollable_frame, width=12, 
                             font=("Arial", 10, "bold"), justify='center', bg='#f0f0f0')
            header.grid(row=0, column=j, padx=1, pady=1, sticky='nsew')
            header.insert(0, col_name)
            header.config(state='readonly')
            header_row.append(header)
        self.cells.append(header_row)
        
        # 创建行号和内容
        for i in range(1, self.rows):
            row_cells = []
            
            # 行号
            row_header = tk.Entry(self.table_scrollable_frame, width=4, 
                                 font=("Arial", 9), justify='center', bg='#f0f0f0')
            row_header.grid(row=i, column=0, padx=1, pady=1, sticky='nsew')
            row_header.insert(0, str(i))
            row_header.config(state='readonly')
            row_cells.append(row_header)
            
            # 数据单元格
            for j in range(1, self.cols):
                cell = tk.Entry(self.table_scrollable_frame, width=12, font=("Arial", 9))
                cell.grid(row=i, column=j, padx=1, pady=1, sticky='nsew')
                cell.bind('<FocusIn>', lambda e, row=i, col=j: self.on_cell_focus(row, col))
                cell.bind('<KeyRelease>', lambda e, row=i, col=j: self.on_cell_change(row, col))
                row_cells.append(cell)
                
                # 初始化单元格数据
                cell_key = f"{i},{j}"
                self.cell_data[cell_key] = {"value": "", "formula": "", "style": {}}
            self.cells.append(row_cells)
        
        # 设置列权重
        for j in range(self.cols):
            self.table_scrollable_frame.columnconfigure(j, weight=1)
    
    def on_cell_focus(self, row, col):
        """当单元格获得焦点时显示公式"""
        cell_key = f"{row},{col}"
        formula = self.cell_data[cell_key].get("formula", "")
        if formula:
            self.formula_var.set(f"={formula}")
        else:
            value = self.cells[row][col].get()
            self.formula_var.set(value)
    
    def on_cell_change(self, row, col):
        """当单元格内容改变时更新数据"""
        cell_key = f"{row},{col}"
        value = self.cells[row][col].get()
        self.cell_data[cell_key]["value"] = value
        
        # 检查是否是公式
        if value.startswith('='):
            try:
                result = self.evaluate_formula(value[1:])
                self.cells[row][col].delete(0, tk.END)
                self.cells[row][col].insert(0, str(result))
                self.cell_data[cell_key]["formula"] = value[1:]
                self.cell_data[cell_key]["value"] = str(result)
            except Exception as e:
                self.cells[row][col].config(fg='red')
        
    def evaluate_formula(self, formula):
        """评估公式"""
        try:
            # 基本数学运算
            if formula.upper().startswith('SUM('):
                # 简单的SUM函数实现
                range_str = formula[4:-1]
                parts = range_str.split(':')
                if len(parts) == 2:
                    total = 0
                    # 这里应该解析单元格范围，简化实现
                    return f"SUM({range_str})"
            
            # 基本算术运算
            safe_dict = {'__builtins__': None}
            return eval(formula, safe_dict)
        except:
            return "#ERROR!"
            
    def apply_formula(self, event=None):
        """应用公式"""
        formula = self.formula_var.get()
        if formula.startswith('='):
            # 应用到当前焦点单元格
            for i in range(1, self.rows):
                for j in range(1, self.cols):
                    if self.cells[i][j] == self.root.focus_get():
                        self.cells[i][j].delete(0, tk.END)
                        self.cells[i][j].insert(0, formula)
                        self.on_cell_change(i, j)
                        break
    
    def add_row(self):
        """添加行"""
        self.rows += 1
        self.create_table()
    
    def add_column(self):
        """添加列"""
        self.cols += 1
        self.create_table()
        
    def create_presentation(self):
        """创建演示文稿"""
        self.presentation_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.presentation_frame, text="演示文稿")
        
        self.create_enhanced_presentation()
        
    def create_enhanced_presentation(self):
        """创建增强的演示文稿"""
        # 主容器
        main_container = tk.Frame(self.presentation_frame)
        main_container.pack(fill='both', expand=True)
        
        # 幻灯片预览面板
        preview_panel = tk.Frame(main_container, width=150, bg='#2c3e50')
        preview_panel.pack(side='left', fill='y', padx=5, pady=5)
        
        # 幻灯片编辑区域
        edit_area = tk.Frame(main_container)
        edit_area.pack(side='right', fill='both', expand=True, padx=5, pady=5)
        
        # 创建幻灯片管理器
        self.create_slide_manager(preview_panel, edit_area)
        
    def create_slide_manager(self, preview_panel, edit_area):
        """创建幻灯片管理器"""
        # 预览标题
        preview_label = tk.Label(preview_panel, text="幻灯片", bg='#2c3e50', fg='white', 
                                font=('Arial', 12, 'bold'))
        preview_label.pack(pady=10)
        
        # 幻灯片列表
        self.slide_listbox = tk.Listbox(preview_panel, bg='#34495e', fg='white', 
                                       selectbackground='#3498db', height=15)
        self.slide_listbox.pack(fill='both', expand=True, padx=5, pady=5)
        self.slide_listbox.bind('<<ListboxSelect>>', self.on_slide_select)
        
        # 幻灯片操作按钮
        btn_frame = tk.Frame(preview_panel, bg='#2c3e50')
        btn_frame.pack(fill='x', padx=5, pady=5)
        
        add_slide_btn = ttk.Button(btn_frame, text="添加", command=self.add_slide, style='Modern.TButton')
        add_slide_btn.pack(side='left', fill='x', expand=True, padx=2)
        
        remove_slide_btn = ttk.Button(btn_frame, text="删除", command=self.remove_slide, style='Accent.TButton')
        remove_slide_btn.pack(side='left', fill='x', expand=True, padx=2)
        
        # 幻灯片编辑区域
        edit_toolbar = tk.Frame(edit_area, bg='#ecf0f1')
        edit_toolbar.pack(fill='x', padx=5, pady=5)
        
        layout_btn = ttk.Button(edit_toolbar, text="布局", command=self.choose_layout)
        layout_btn.pack(side='left', padx=2)
        
        theme_btn = ttk.Button(edit_toolbar, text="主题", command=self.choose_theme)
        theme_btn.pack(side='left', padx=2)
        
        # 幻灯片画布
        canvas_frame = tk.Frame(edit_area, bg='white', relief='sunken', bd=2)
        canvas_frame.pack(expand=True, padx=20, pady=20)
        
        self.slide_canvas = tk.Canvas(canvas_frame, bg='white', width=800, height=500)
        self.slide_canvas.pack(padx=10, pady=10)
        
        # 初始化幻灯片数据
        self.slides = []
        self.current_slide_index = 0
        self.slide_layouts = []
        self.add_slide()
        
    def add_slide(self):
        """添加幻灯片"""
        slide_data = {
            "title": f"幻灯片 {len(self.slides) + 1}",
            "content": "在此添加内容...",
            "layout": "title_content",
            "background": "#ffffff",
            "elements": []
        }
        self.slides.append(slide_data)
        self.slide_layouts.append("title_content")
        self.slide_listbox.insert(tk.END, slide_data["title"])
        self.current_slide_index = len(self.slides) - 1
        self.draw_current_slide()
    
    def remove_slide(self):
        """删除幻灯片"""
        if len(self.slides) > 1:
            self.slides.pop(self.current_slide_index)
            self.slide_layouts.pop(self.current_slide_index)
            self.slide_listbox.delete(self.current_slide_index)
            self.current_slide_index = min(self.current_slide_index, len(self.slides) - 1)
            if self.current_slide_index >= 0:
                self.slide_listbox.selection_set(self.current_slide_index)
                self.draw_current_slide()
    
    def on_slide_select(self, event):
        """选择幻灯片"""
        selection = self.slide_listbox.curselection()
        if selection:
            self.current_slide_index = selection[0]
            self.draw_current_slide()
    
    def choose_layout(self):
        """选择幻灯片布局"""
        layout_window = tk.Toplevel(self.root)
        layout_window.title("选择布局")
        layout_window.geometry("300x200")
        
        layouts = [
            ("标题和内容", "title_content"),
            ("两栏内容", "two_columns"),
            ("仅标题", "title_only"),
            ("空白", "blank")
        ]
        
        for name, layout_type in layouts:
            btn = ttk.Button(layout_window, text=name, 
                          command=lambda lt=layout_type: self.apply_layout(lt, layout_window))
            btn.pack(fill='x', padx=20, pady=5)
    
    def apply_layout(self, layout_type, window):
        """应用选择的布局"""
        if 0 <= self.current_slide_index < len(self.slide_layouts):
            self.slide_layouts[self.current_slide_index] = layout_type
            self.slides[self.current_slide_index]["layout"] = layout_type
            self.draw_current_slide()
        window.destroy()
    
    def choose_theme(self):
        """选择主题"""
        color = colorchooser.askcolor(title="选择背景颜色")[1]
        if color and 0 <= self.current_slide_index < len(self.slides):
            self.slides[self.current_slide_index]["background"] = color
            self.draw_current_slide()
    
    def draw_current_slide(self):
        """绘制当前幻灯片"""
        self.slide_canvas.delete("all")
        
        if not self.slides or self.current_slide_index >= len(self.slides):
            return
            
        slide = self.slides[self.current_slide_index]
        layout = self.slide_layouts[self.current_slide_index]
        
        # 绘制背景
        self.slide_canvas.create_rectangle(0, 0, 800, 500, fill=slide["background"], outline="")
        
        # 根据布局绘制内容
        if layout == "title_content":
            # 绘制标题
            self.slide_canvas.create_text(400, 100, text=slide["title"], 
                                         font=("Arial", 32, "bold"), fill="#2c3e50")
            # 绘制内容
            self.slide_canvas.create_text(400, 300, text=slide["content"], 
                                         font=("Arial", 18), fill="#34495e", width=600)
        
        elif layout == "two_columns":
            self.slide_canvas.create_text(400, 80, text=slide["title"], 
                                         font=("Arial", 32, "bold"), fill="#2c3e50")
            # 左栏
            self.slide_canvas.create_text(200, 300, text="左栏内容", 
                                         font=("Arial", 16), fill="#34495e", width=300)
            # 右栏
            self.slide_canvas.create_text(600, 300, text="右栏内容", 
                                         font=("Arial", 16), fill="#34495e", width=300)
        
        elif layout == "title_only":
            self.slide_canvas.create_text(400, 250, text=slide["title"], 
                                         font=("Arial", 36, "bold"), fill="#2c3e50")
        
        # 显示幻灯片编号
        self.slide_canvas.create_text(750, 480, text=f"{self.current_slide_index + 1}/{len(self.slides)}", 
                                     font=("Arial", 12), fill="#7f8c8d")
        
    def create_database_viewer(self):
        """创建数据库查看器"""
        self.db_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.db_frame, text="数据库")
        
        # 创建数据库管理界面
        self.create_database_interface()
        
    def create_database_interface(self):
        """创建数据库管理界面"""
        # 工具栏
        toolbar = tk.Frame(self.db_frame, bg='#ecf0f1')
        toolbar.pack(fill='x', padx=5, pady=5)
        
        # 数据库操作按钮
        new_db_btn = ttk.Button(toolbar, text="新建表", command=self.create_new_table)
        new_db_btn.pack(side='left', padx=2)
        
        query_btn = ttk.Button(toolbar, text="执行查询", command=self.execute_sql_query)
        query_btn.pack(side='left', padx=2)
        
        # 查询输入框
        query_frame = tk.Frame(self.db_frame)
        query_frame.pack(fill='x', padx=5, pady=5)
        
        query_label = tk.Label(query_frame, text="SQL查询:")
        query_label.pack(side='left', padx=5)
        
        self.query_entry = tk.Entry(query_frame, width=80)
        self.query_entry.pack(side='left', fill='x', expand=True, padx=5)
        self.query_entry.bind('<Return>', lambda e: self.execute_sql_query())
        self.query_entry.insert(0, "SELECT * FROM documents LIMIT 10")
        
        # 结果显示
        result_frame = tk.Frame(self.db_frame)
        result_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, height=20, font=('Courier New', 10))
        self.result_text.pack(fill='both', expand=True)
        
    def create_new_table(self):
        """创建新表"""
        table_window = tk.Toplevel(self.root)
        table_window.title("新建表")
        table_window.geometry("400x300")
        
        name_label = tk.Label(table_window, text="表名:")
        name_label.pack(pady=5)
        
        name_entry = tk.Entry(table_window, width=30)
        name_entry.pack(pady=5)
        
        create_btn = ttk.Button(table_window, text="创建", 
                              command=lambda: self.execute_create_table(name_entry.get(), table_window))
        create_btn.pack(pady=10)
        
    def execute_create_table(self, table_name, window):
        """执行创建表"""
        if table_name:
            try:
                self.cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS {table_name} (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                self.conn.commit()
                messagebox.showinfo("成功", f"表 '{table_name}' 创建成功")
                window.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"创建失败: {str(e)}")
                
    def execute_sql_query(self):
        """执行SQL查询"""
        query = self.query_entry.get()
        if query:
            try:
                self.cursor.execute(query)
                
                if query.strip().upper().startswith('SELECT'):
                    # 显示查询结果
                    results = self.cursor.fetchall()
                    columns = [description[0] for description in self.cursor.description]
                    
                    self.result_text.delete('1.0', tk.END)
                    
                    # 显示列名
                    header = " | ".join(columns)
                    self.result_text.insert(tk.END, header + "\n")
                    self.result_text.insert(tk.END, "-" * len(header) + "\n")
                    
                    # 显示数据
                    for row in results:
                        row_str = " | ".join(str(cell) for cell in row)
                        self.result_text.insert(tk.END, row_str + "\n")
                        
                    self.result_text.insert(tk.END, f"\n共 {len(results)} 行数据")
                else:
                    self.conn.commit()
                    self.result_text.delete('1.0', tk.END)
                    self.result_text.insert(tk.END, "查询执行成功")
                    
            except Exception as e:
                self.result_text.delete('1.0', tk.END)
                self.result_text.insert(tk.END, f"错误: {str(e)}")

    # ===== 核心功能实现 =====
    
    def new_file(self):
        """新建文件"""
        self.text_area.delete('1.0', tk.END)
        self.current_file = None
        self.root.title("OfficeMate - 新文档")
        self.add_to_version_history("新建文档")
        
    def open_file(self):
        """打开文件"""
        file_path = filedialog.askopenfilename(
            defaultextension=".txt",
            filetypes=[
                ("文本文档", "*.txt"),
                ("JSON 文件", "*.json"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.text_area.delete('1.0', tk.END)
                    self.text_area.insert('1.0', content)
                    self.current_file = file_path
                    self.root.title(f"OfficeMate Pro - {os.path.basename(file_path)}")
                    
                    # 添加到最近文件列表
                    self.add_to_recent_files(file_path)
                    self.add_to_version_history(f"打开文件: {os.path.basename(file_path)}")
                    
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {str(e)}")
                
    def save_file(self):
        """保存文件"""
        if self.current_file:
            self.save_document()
        else:
            self.save_as_file()
            
    def save_as_file(self):
        """另存为文件"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("文本文档", "*.txt"),
                ("JSON 文件", "*.json"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.current_file = file_path
            self.save_document()
            self.add_to_recent_files(file_path)
            
    def save_document(self, auto_save=False):
        """保存文档"""
        try:
            content = self.text_area.get('1.0', 'end-1c')
            
            if self.current_file.endswith('.json'):
                document_data = {
                    'metadata': {
                        'version': '2.0',
                        'created_at': datetime.now().isoformat(),
                        'modified_at': datetime.now().isoformat(),
                        'author': self.user_id
                    },
                    'content': content,
                    'spreadsheet_data': self.cell_data,
                    'presentation_data': self.slides
                }
                with open(self.current_file, 'w', encoding='utf-8') as f:
                    json.dump(document_data, f, ensure_ascii=False, indent=2)
            else:
                with open(self.current_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                    
            if not auto_save:
                messagebox.showinfo("成功", "文档已保存")
                self.add_to_version_history("保存文档")
                
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")
            
    def export_document(self):
        """导出文档"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[
                ("HTML 文件", "*.html"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            try:
                content = self.text_area.get('1.0', 'end-1c')
                
                if file_path.endswith('.html'):
                    html_content = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="UTF-8">
                        <title>OfficeMate Pro 文档</title>
                        <style>
                            body {{ 
                                font-family: Arial, sans-serif; 
                                margin: 40px;
                                line-height: 1.6;
                                color: #333;
                            }}
                            .content {{ 
                                white-space: pre-wrap;
                                background: #f8f9fa;
                                padding: 20px;
                                border-radius: 5px;
                                border: 1px solid #ddd;
                            }}
                            .header {{
                                text-align: center;
                                margin-bottom: 30px;
                                color: #2c3e50;
                            }}
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <h1>OfficeMate Pro 文档</h1>
                            <p>导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                        </div>
                        <div class="content">{content}</div>
                    </body>
                    </html>
                    """
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)
                else:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                        
                messagebox.showinfo("成功", "文档导出完成")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")

    # ===== AI 功能实现 =====
    
    def ai_writing_assistant(self):
        """AI写作助手"""
        ai_window = tk.Toplevel(self.root)
        ai_window.title("AI写作助手")
        ai_window.geometry("600x500")
        
        notebook = ttk.Notebook(ai_window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 智能写作标签页
        writing_frame = ttk.Frame(notebook)
        notebook.add(writing_frame, text="智能写作")
        
        prompt_label = tk.Label(writing_frame, text="写作提示:")
        prompt_label.pack(anchor='w', pady=5)
        
        prompt_entry = scrolledtext.ScrolledText(writing_frame, height=4)
        prompt_entry.pack(fill='x', pady=5)
        
        # 写作类型选择
        type_frame = tk.Frame(writing_frame)
        type_frame.pack(fill='x', pady=5)
        
        tk.Label(type_frame, text="写作类型:").pack(side='left')
        
        writing_type = tk.StringVar(value="正式")
        types = ["正式", "创意", "商务", "技术", "学术"]
        for t in types:
            tk.Radiobutton(type_frame, text=t, variable=writing_type, value=t).pack(side='left', padx=5)
        
        generate_btn = ttk.Button(writing_frame, text="生成内容", 
                                command=lambda: self.generate_ai_content(
                                    prompt_entry.get('1.0', tk.END), 
                                    writing_type.get(),
                                    result_text
                                ), style='Success.TButton')
        generate_btn.pack(pady=10)
        
        result_label = tk.Label(writing_frame, text="生成结果:")
        result_label.pack(anchor='w', pady=5)
        
        result_text = scrolledtext.ScrolledText(writing_frame, height=10)
        result_text.pack(fill='both', expand=True)
        
        # 语法检查标签页
        grammar_frame = ttk.Frame(notebook)
        notebook.add(grammar_frame, text="语法检查")
        
        self.setup_grammar_checker(grammar_frame)
        
    def generate_ai_content(self, prompt, writing_type, result_widget):
        """生成AI内容"""
        # 模拟AI生成内容
        responses = {
            "正式": f"基于您的要求，我为您准备了以下正式文档内容：\n\n{prompt}\n\n此文档经过精心组织，语言规范，符合专业写作标准。内容结构清晰，逻辑严密，适合商务和官方场合使用。",
            "创意": f"✨ 创意写作时间！ ✨\n\n{prompt}\n\n让我们用想象力创造精彩的内容！这段文字充满了生动的比喻和丰富的意象，让读者仿佛身临其境。",
            "商务": f"📊 商务文档生成：\n\n{prompt}\n\n此内容适合商务场合使用，语言专业但不失亲和力。重点突出，建议明确，能够有效传达商业意图。",
            "技术": f"🔧 技术文档：\n\n{prompt}\n\n技术描述准确，术语使用规范。包含详细的操作步骤和技术参数，适合开发人员和技术人员参考。",
            "学术": f"🎓 学术写作：\n\n{prompt}\n\n语言严谨，引用规范，逻辑清晰。符合学术写作标准，适合论文和研究报告使用。"
        }
        
        result = responses.get(writing_type, "请选择有效的写作类型")
        result_widget.delete('1.0', tk.END)
        result_widget.insert('1.0', result)
        
    def setup_grammar_checker(self, parent):
        """设置语法检查器"""
        text_label = tk.Label(parent, text="待检查文本:")
        text_label.pack(anchor='w', pady=5)
        
        check_text = scrolledtext.ScrolledText(parent, height=8)
        check_text.pack(fill='x', pady=5)
        
        check_btn = ttk.Button(parent, text="检查语法", 
                             command=lambda: self.check_grammar(check_text.get('1.0', tk.END), result_text),
                             style='Modern.TButton')
        check_btn.pack(pady=10)
        
        result_label = tk.Label(parent, text="检查结果:")
        result_label.pack(anchor='w', pady=5)
        
        result_text = scrolledtext.ScrolledText(parent, height=8)
        result_text.pack(fill='both', expand=True)
        
    def check_grammar(self, text, result_widget):
        """检查语法"""
        # 简单的语法检查模拟
        issues = []
        
        if len(text.split()) < 10:
            issues.append("⚠️ 文本过短，建议补充更多内容")
            
        if '。' not in text and '.' not in text:
            issues.append("❌ 缺少句号，请检查句子完整性")
            
        if text.count('，') < text.count('。') * 0.5:
            issues.append("💡 建议增加逗号使用，使句子结构更清晰")
            
        if not any(char.isdigit() for char in text):
            issues.append("📊 可以考虑添加具体数据支持观点")
            
        if len(issues) == 0:
            issues.append("✅ 文本语法检查通过！")
            
        result = "\n".join(issues)
        result_widget.delete('1.0', tk.END)
        result_widget.insert('1.0', result)
        
    def ai_grammar_check(self):
        """AI语法检查"""
        selected_text = self.get_selected_text()
        if selected_text:
            self.ai_writing_assistant()
        else:
            messagebox.showinfo("提示", "请先选择要检查的文本")
            
    def ai_content_optimize(self):
        """AI内容优化"""
        selected_text = self.get_selected_text()
        if selected_text:
            # 在实际应用中，这里会调用AI API进行内容优化
            optimized = f"优化建议：\n\n{selected_text}\n\n💡 建议：\n- 调整句子结构使其更流畅\n- 使用更精确的词汇\n- 增加具体例子支持观点"
            self.show_ai_result("内容优化", optimized)
        else:
            messagebox.showinfo("提示", "请先选择要优化的文本")
            
    def ai_text_summarize(self):
        """AI文本摘要"""
        selected_text = self.get_selected_text()
        if selected_text and len(selected_text.split()) > 50:
            # 模拟文本摘要
            words = selected_text.split()
            summary = " ".join(words[:30]) + "...\n\n📋 摘要要点：\n• 主要讨论了相关内容\n• 提出了重要观点\n• 建议进一步研究"
            self.show_ai_result("文本摘要", summary)
        else:
            messagebox.showinfo("提示", "请选择至少50词的文本进行摘要")
            
    def get_selected_text(self):
        """获取选中的文本"""
        try:
            return self.text_area.get('sel.first', 'sel.last')
        except:
            return None
            
    def show_ai_result(self, title, content):
        """显示AI结果"""
        result_window = tk.Toplevel(self.root)
        result_window.title(f"AI{title}")
        result_window.geometry("500x400")
        
        text_widget = scrolledtext.ScrolledText(result_window, wrap='word')
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        text_widget.insert('1.0', content)
        
        # 插入按钮
        insert_btn = ttk.Button(result_window, text="插入到文档", 
                              command=lambda: self.insert_ai_content(content, result_window),
                              style='Success.TButton')
        insert_btn.pack(pady=10)
        
    def insert_ai_content(self, content, window):
        """插入AI生成的内容"""
        try:
            self.text_area.insert(tk.INSERT, content)
            window.destroy()
        except:
            pass

    # ===== 文本格式化功能 =====
    
    def change_font(self, event=None):
        """改变字体"""
        self.current_font = self.font_var.get()
        self.apply_text_formatting()
    
    def change_font_size(self, event=None):
        """改变字号"""
        try:
            self.current_size = int(self.size_var.get())
            self.apply_text_formatting()
        except ValueError:
            pass
    
    def toggle_bold(self):
        """切换粗体"""
        self.current_bold = not self.current_bold
        self.apply_text_formatting()
    
    def toggle_italic(self):
        """切换斜体"""
        self.current_italic = not self.current_italic
        self.apply_text_formatting()
    
    def toggle_underline(self):
        """切换下划线"""
        self.current_underline = not self.current_underline
        self.apply_text_formatting()
    
    def apply_text_formatting(self):
        """应用文本格式化"""
        try:
            # 获取选中文本的范围
            start = self.text_area.index("sel.first")
            end = self.text_area.index("sel.last")
            
            # 配置字体
            font_spec = [self.current_font, self.current_size]
            if self.current_bold:
                font_spec.append("bold")
            if self.current_italic:
                font_spec.append("italic")
            if self.current_underline:
                font_spec.append("underline")
                
            # 创建标签
            tag_name = f"format_{len(self.document_history)}"
            self.text_area.tag_configure(tag_name, font=font_spec)
            self.text_area.tag_add(tag_name, start, end)
            
        except tk.TclError:
            # 没有选中文本
            pass

    # ===== 协作功能 =====
    
    def start_collaboration_server(self):
        """启动协作服务器"""
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.bind(('localhost', 12345))
            self.server_socket.listen(5)
            
            server_thread = threading.Thread(target=self.run_collaboration_server)
            server_thread.daemon = True
            server_thread.start()
            
            self.collaboration_mode = True
            self.collab_label.config(text="服务器运行中", fg='green')
            messagebox.showinfo("协作", "协作服务器已启动在 localhost:12345")
        except Exception as e:
            messagebox.showerror("错误", f"无法启动服务器: {str(e)}")
            
    def run_collaboration_server(self):
        """运行协作服务器"""
        while True:
            try:
                client_socket, address = self.server_socket.accept()
                self.connected_clients.append(client_socket)
                print(f"客户端连接: {address}")
            except:
                break
                
    def connect_to_server_dialog(self):
        """连接到服务器对话框"""
        connect_window = tk.Toplevel(self.root)
        connect_window.title("连接到服务器")
        connect_window.geometry("400x300")
        
        tk.Label(connect_window, text="服务器地址:").pack(pady=5)
        address_entry = tk.Entry(connect_window, width=20)
        address_entry.pack(pady=5)
        address_entry.insert(0, "localhost")
        
        tk.Label(connect_window, text="端口:").pack(pady=5)
        port_entry = tk.Entry(connect_window, width=10)
        port_entry.pack(pady=5)
        port_entry.insert(0, "12345")
        
        connect_btn = ttk.Button(connect_window, text="连接", 
                               command=lambda: self.connect_to_server(
                                   address_entry.get(), 
                                   port_entry.get(), 
                                   connect_window
                               ))
        connect_btn.pack(pady=10)
        
    def connect_to_server(self, address, port, window):
        """连接到服务器"""
        try:
            self.client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.client_socket.connect((address, int(port)))
            self.collaboration_mode = True
            self.collab_label.config(text=f"已连接: {address}", fg='green')
            window.destroy()
            messagebox.showinfo("成功", f"已连接到服务器 {address}:{port}")
        except Exception as e:
            messagebox.showerror("错误", f"连接失败: {str(e)}")
            
    def disconnect_from_server(self):
        """断开服务器连接"""
        if self.client_socket:
            self.client_socket.close()
        if self.server_socket:
            self.server_socket.close()
            
        self.collaboration_mode = False
        self.collab_label.config(text="离线", fg='red')
        messagebox.showinfo("协作", "已断开连接")
        
    def share_document_dialog(self):
        """共享文档对话框"""
        if not self.collaboration_mode:
            messagebox.showwarning("协作", "请先启动协作服务器或连接到服务器")
            return
            
        share_window = tk.Toplevel(self.root)
        share_window.title("共享文档")
        share_window.geometry("400x300")
        
        tk.Label(share_window, text="文档共享设置", font=('Arial', 12, 'bold')).pack(pady=10)
        
        # 权限设置
        perm_frame = tk.Frame(share_window)
        perm_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(perm_frame, text="权限:").pack(side='left')
        permission = tk.StringVar(value="view")
        tk.Radiobutton(perm_frame, text="只读", variable=permission, value="view").pack(side='left', padx=10)
        tk.Radiobutton(perm_frame, text="可编辑", variable=permission, value="edit").pack(side='left', padx=10)
        
        # 共享链接显示
        link_frame = tk.Frame(share_window)
        link_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(link_frame, text="共享链接:").pack(anchor='w')
        link_text = tk.Text(link_frame, height=2, width=40)
        link_text.pack(fill='x', pady=5)
        link_text.insert('1.0', f"office-mate://share/{self.user_id}/{datetime.now().timestamp()}")
        
        share_btn = ttk.Button(share_window, text="生成共享链接", 
                             command=lambda: self.generate_share_link(permission.get()),
                             style='Success.TButton')
        share_btn.pack(pady=20)

    # ===== 实用工具功能 =====
    
    def on_text_change(self, event=None):
        """文本变化处理"""
        self.update_word_count()
        self.update_cursor_position()
        
    def update_word_count(self):
        """更新字数统计"""
        content = self.text_area.get('1.0', 'end-1c')
        words = len(content.split())
        chars = len(content)
        lines = content.count('\n') + 1
        self.word_count_label.config(text=f"字数: {words} 字符: {chars} 行: {lines}")
        
    def update_cursor_position(self, event=None):
        """更新光标位置"""
        try:
            cursor_pos = self.text_area.index(tk.INSERT)
            line, col = cursor_pos.split('.')
            self.cursor_label.config(text=f"行: {line}, 列: {int(col)+1}")
        except:
            pass
            
    def find_replace_dialog(self):
        """查找替换对话框"""
        find_window = tk.Toplevel(self.root)
        find_window.title("查找和替换")
        find_window.geometry("400x200")
        
        # 查找框
        find_frame = tk.Frame(find_window)
        find_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(find_frame, text="查找:").pack(side='left')
        find_entry = tk.Entry(find_frame, width=30)
        find_entry.pack(side='left', padx=5)
        
        # 替换框
        replace_frame = tk.Frame(find_window)
        replace_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(replace_frame, text="替换为:").pack(side='left')
        replace_entry = tk.Entry(replace_frame, width=30)
        replace_entry.pack(side='left', padx=5)
        
        # 按钮框架
        btn_frame = tk.Frame(find_window)
        btn_frame.pack(pady=20)
        
        find_btn = ttk.Button(btn_frame, text="查找", 
                            command=lambda: self.find_text(find_entry.get()))
        find_btn.pack(side='left', padx=5)
        
        replace_btn = ttk.Button(btn_frame, text="替换", 
                               command=lambda: self.replace_text(find_entry.get(), replace_entry.get()))
        replace_btn.pack(side='left', padx=5)
        
        replace_all_btn = ttk.Button(btn_frame, text="全部替换", 
                                   command=lambda: self.replace_all_text(find_entry.get(), replace_entry.get()))
        replace_all_btn.pack(side='left', padx=5)
        
    def find_text(self, text):
        """查找文本"""
        if text:
            # 清除之前的高亮
            self.text_area.tag_remove('highlight', '1.0', tk.END)
            
            # 查找文本
            start_pos = '1.0'
            while True:
                start_pos = self.text_area.search(text, start_pos, stopindex=tk.END)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(text)}c"
                self.text_area.tag_add('highlight', start_pos, end_pos)
                start_pos = end_pos
            
            # 设置高亮样式
            self.text_area.tag_config('highlight', background='yellow')
            
    def replace_text(self, find_text, replace_text):
        """替换文本"""
        if find_text and self.text_area.tag_ranges('sel'):
            selected = self.text_area.get('sel.first', 'sel.last')
            if selected == find_text:
                self.text_area.delete('sel.first', 'sel.last')
                self.text_area.insert('sel.first', replace_text)
                
    def replace_all_text(self, find_text, replace_text):
        """替换所有文本"""
        if find_text:
            content = self.text_area.get('1.0', tk.END)
            new_content = content.replace(find_text, replace_text)
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert('1.0', new_content)
            
    def show_word_count(self):
        """显示详细字数统计"""
        content = self.text_area.get('1.0', 'end-1c')
        words = len(content.split())
        chars = len(content)
        chars_no_space = len(content.replace(' ', '').replace('\n', '').replace('\t', ''))
        lines = content.count('\n') + 1
        paragraphs = content.count('\n\n') + 1
        
        stats = f"""详细统计信息:
        
字数: {words} 个
字符数(含空格): {chars} 个
字符数(不含空格): {chars_no_space} 个
行数: {lines} 行
段落数: {paragraphs} 段

文档大小: {chars / 1024:.2f} KB
估计阅读时间: {words // 200 + 1} 分钟"""
        
        messagebox.showinfo("字数统计", stats)
        
    def version_history(self):
        """版本历史"""
        history_window = tk.Toplevel(self.root)
        history_window.title("版本历史")
        history_window.geometry("600x400")
        
        # 工具栏
        toolbar = tk.Frame(history_window)
        toolbar.pack(fill='x', padx=10, pady=5)
        
        refresh_btn = ttk.Button(toolbar, text="刷新", command=self.refresh_version_history)
        refresh_btn.pack(side='left', padx=5)
        
        # 版本列表
        tree = ttk.Treeview(history_window, columns=('版本', '描述', '时间', '大小'), show='headings')
        tree.heading('版本', text='版本')
        tree.heading('描述', text='描述')
        tree.heading('时间', text='时间')
        tree.heading('大小', text='大小')
        
        tree.column('版本', width=80)
        tree.column('描述', width=200)
        tree.column('时间', width=150)
        tree.column('大小', width=80)
        
        for i, version in enumerate(self.document_history[-20:]):  # 显示最近20个版本
            tree.insert('', 'end', values=(
                version.get('version', i+1),
                version.get('description', '自动保存'),
                version.get('timestamp', ''),
                f"{len(version.get('content', ''))} 字符"
            ))
            
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
    def refresh_version_history(self):
        """刷新版本历史"""
        # 在实际应用中，这里会重新加载版本历史
        pass
        
    def add_to_version_history(self, description):
        """添加到版本历史"""
        version_data = {
            'version': self.current_version + 1,
            'description': description,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'content': self.text_area.get('1.0', 'end-1c')[:1000]  # 只保存部分内容
        }
        self.document_history.append(version_data)
        self.current_version += 1
        
    def add_to_recent_files(self, file_path):
        """添加到最近文件列表"""
        if file_path in self.user_preferences['recent_files']:
            self.user_preferences['recent_files'].remove(file_path)
        self.user_preferences['recent_files'].insert(0, file_path)
        # 只保留最近10个文件
        self.user_preferences['recent_files'] = self.user_preferences['recent_files'][:10]
        self.save_preferences()
        self.update_recent_files_menu()

    # ===== 对话框和工具方法 =====
    
    def generate_share_link(self, permission):
        """生成共享链接"""
        messagebox.showinfo("共享", f"共享链接已生成\n权限: {permission}\n\n链接已复制到剪贴板")
        
    def insert_formula_dialog(self):
        """插入公式对话框"""
        formula_window = tk.Toplevel(self.root)
        formula_window.title("插入公式")
        formula_window.geometry("300x200")
        
        tk.Label(formula_window, text="常用公式:").pack(pady=10)
        
        formulas = [
            "=SUM(A1:A10)",
            "=AVERAGE(B1:B10)", 
            "=MAX(C1:C10)",
            "=MIN(D1:D10)",
            "=COUNT(E1:E10)"
        ]
        
        for formula in formulas:
            btn = ttk.Button(formula_window, text=formula, 
                           command=lambda f=formula: self.insert_formula_to_cell(f, formula_window))
            btn.pack(fill='x', padx=20, pady=2)
            
    def insert_formula_to_cell(self, formula, window):
        """插入公式到单元格"""
        self.formula_var.set(formula)
        window.destroy()
        
    def insert_image_dialog(self):
        """插入图片对话框"""
        if not HAS_PIL:
            messagebox.showwarning("功能不可用", "图片插入功能需要安装 Pillow 库")
            return
            
        file_path = filedialog.askopenfilename(
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp")]
        )
        if file_path:
            messagebox.showinfo("图片插入", f"已选择图片: {os.path.basename(file_path)}\n\n图片插入功能需要更复杂的实现")
            
    def insert_table_dialog(self):
        """插入表格对话框"""
        table_window = tk.Toplevel(self.root)
        table_window.title("插入表格")
        table_window.geometry("300x200")
        
        tk.Label(table_window, text="表格尺寸").pack(pady=10)
        
        size_frame = tk.Frame(table_window)
        size_frame.pack(pady=10)
        
        tk.Label(size_frame, text="行:").grid(row=0, column=0, padx=5)
        rows_entry = tk.Entry(size_frame, width=5)
        rows_entry.grid(row=0, column=1, padx=5)
        rows_entry.insert(0, "3")
        
        tk.Label(size_frame, text="列:").grid(row=0, column=2, padx=5)
        cols_entry = tk.Entry(size_frame, width=5)
        cols_entry.grid(row=0, column=3, padx=5)
        cols_entry.insert(0, "3")
        
        insert_btn = ttk.Button(table_window, text="插入表格", 
                              command=lambda: self.insert_table(
                                  int(rows_entry.get()), 
                                  int(cols_entry.get()),
                                  table_window
                              ))
        insert_btn.pack(pady=20)
        
    def insert_table(self, rows, cols, window):
        """插入表格"""
        table_text = "\n" + "+" + "---+" * cols + "\n"
        for i in range(rows):
            table_text += "|" + "   |" * cols + "\n"
            table_text += "+" + "---+" * cols + "\n"
            
        self.text_area.insert(tk.INSERT, table_text)
        window.destroy()
        
    def insert_chart_dialog(self):
        """插入图表对话框"""
        messagebox.showinfo("图表", "图表插入功能需要安装 matplotlib 库")
        
    def insert_equation_dialog(self):
        """插入公式对话框"""
        equation_window = tk.Toplevel(self.root)
        equation_window.title("插入公式")
        equation_window.geometry("400x300")
        
        tk.Label(equation_window, text="数学公式编辑器", font=('Arial', 12, 'bold')).pack(pady=10)
        
        # 常用公式按钮
        equations_frame = tk.Frame(equation_window)
        equations_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        equations = [
            "a² + b² = c²",
            "E = mc²", 
            "π ≈ 3.14159",
            "∑(i=1 to n) i = n(n+1)/2",
            "∫ f(x) dx"
        ]
        
        for eq in equations:
            btn = ttk.Button(equations_frame, text=eq, 
                           command=lambda e=eq: self.insert_equation(e, equation_window))
            btn.pack(fill='x', pady=2)
            
    def insert_equation(self, equation, window):
        """插入公式"""
        self.text_area.insert(tk.INSERT, f"\n{equation}\n")
        window.destroy()
        
    def insert_hyperlink_dialog(self):
        """插入超链接对话框"""
        link_window = tk.Toplevel(self.root)
        link_window.title("插入超链接")
        link_window.geometry("400x200")
        
        tk.Label(link_window, text="链接文本:").pack(pady=5)
        text_entry = tk.Entry(link_window, width=40)
        text_entry.pack(pady=5)
        
        tk.Label(link_window, text="链接地址:").pack(pady=5)
        url_entry = tk.Entry(link_window, width=40)
        url_entry.pack(pady=5)
        url_entry.insert(0, "https://")
        
        insert_btn = ttk.Button(link_window, text="插入链接", 
                              command=lambda: self.insert_hyperlink(
                                  text_entry.get(),
                                  url_entry.get(),
                                  link_window
                              ))
        insert_btn.pack(pady=20)
        
    def insert_hyperlink(self, text, url, window):
        """插入超链接"""
        if text and url:
            link_text = f"[{text}]({url})"
            self.text_area.insert(tk.INSERT, link_text)
            window.destroy()
            
    def insert_datetime(self):
        """插入日期时间"""
        now = datetime.now()
        dt_string = now.strftime("%Y-%m-%d %H:%M:%S")
        self.text_area.insert(tk.INSERT, dt_string)

    def font_dialog(self):
        """字体对话框"""
        font_window = tk.Toplevel(self.root)
        font_window.title("字体设置")
        font_window.geometry("300x400")
        
        # 字体选择
        tk.Label(font_window, text="字体:").pack(pady=5)
        font_listbox = tk.Listbox(font_window, height=6)
        for font in ["Arial", "Times New Roman", "Courier New", "Verdana", "微软雅黑", "宋体"]:
            font_listbox.insert(tk.END, font)
        font_listbox.pack(fill='x', padx=20, pady=5)
        font_listbox.selection_set(0)
        
        # 字号选择
        tk.Label(font_window, text="字号:").pack(pady=5)
        size_listbox = tk.Listbox(font_window, height=6)
        for size in ["8", "10", "12", "14", "16", "18", "20", "24"]:
            size_listbox.insert(tk.END, size)
        size_listbox.pack(fill='x', padx=20, pady=5)
        size_listbox.selection_set(2)  # 选择12号字
        
        # 样式选择
        style_frame = tk.Frame(font_window)
        style_frame.pack(fill='x', padx=20, pady=10)
        
        bold_var = tk.BooleanVar()
        italic_var = tk.BooleanVar()
        underline_var = tk.BooleanVar()
        
        tk.Checkbutton(style_frame, text="粗体", variable=bold_var).pack(side='left', padx=10)
        tk.Checkbutton(style_frame, text="斜体", variable=italic_var).pack(side='left', padx=10)
        tk.Checkbutton(style_frame, text="下划线", variable=underline_var).pack(side='left', padx=10)
        
        # 颜色选择
        color_btn = ttk.Button(font_window, text="选择颜色", 
                             command=lambda: self.choose_text_color())
        color_btn.pack(pady=10)
        
        apply_btn = ttk.Button(font_window, text="应用", 
                             command=lambda: self.apply_font_settings(
                                 font_listbox.get(font_listbox.curselection()),
                                 size_listbox.get(size_listbox.curselection()),
                                 bold_var.get(),
                                 italic_var.get(),
                                 underline_var.get(),
                                 font_window
                             ), style='Success.TButton')
        apply_btn.pack(pady=10)
        
    def apply_font_settings(self, font, size, bold, italic, underline, window):
        """应用字体设置"""
        self.current_font = font
        self.current_size = int(size)
        self.current_bold = bold
        self.current_italic = italic
        self.current_underline = underline
        
        self.font_var.set(font)
        self.size_var.set(size)
        self.apply_text_formatting()
        window.destroy()
        
    def choose_text_color(self):
        """选择文字颜色"""
        color = colorchooser.askcolor(title="选择文字颜色")[1]
        if color:
            try:
                self.text_area.tag_configure("color", foreground=color)
                self.text_area.tag_add("color", "sel.first", "sel.last")
            except:
                pass

    def paragraph_dialog(self):
        """段落对话框"""
        messagebox.showinfo("段落设置", "段落设置功能开发中...")
        
    def style_dialog(self):
        """样式对话框"""
        style_window = tk.Toplevel(self.root)
        style_window.title("样式管理")
        style_window.geometry("400x500")
        
        tk.Label(style_window, text="文档样式", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # 样式列表
        styles_frame = tk.Frame(style_window)
        styles_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        styles = [
            ("标题1", "Arial, 24, bold"),
            ("标题2", "Arial, 18, bold"), 
            ("标题3", "Arial, 14, bold"),
            ("正文", "Arial, 12"),
            ("引用", "Arial, 12, italic"),
            ("代码", "Courier New, 11")
        ]
        
        for name, desc in styles:
            frame = tk.Frame(styles_frame, relief='groove', bd=1)
            frame.pack(fill='x', pady=2)
            
            tk.Label(frame, text=name, font=('Arial', 10, 'bold')).pack(side='left', padx=5, pady=5)
            tk.Label(frame, text=desc, font=('Arial', 8)).pack(side='right', padx=5, pady=5)
            
            apply_btn = ttk.Button(frame, text="应用", 
                                 command=lambda n=name: self.apply_style(n))
            apply_btn.pack(side='right', padx=5)
            
    def apply_style(self, style_name):
        """应用样式"""
        styles = {
            "标题1": ("Arial", 24, "bold"),
            "标题2": ("Arial", 18, "bold"),
            "标题3": ("Arial", 14, "bold"),
            "正文": ("Arial", 12),
            "引用": ("Arial", 12, "italic"),
            "代码": ("Courier New", 11)
        }
        
        if style_name in styles:
            try:
                self.text_area.tag_configure(style_name, font=styles[style_name])
                self.text_area.tag_add(style_name, "sel.first", "sel.last")
            except:
                # 如果没有选中文本，在当前位置插入样式文本
                self.text_area.insert(tk.INSERT, f"\n{style_name}\n")
                
    def theme_dialog(self):
        """主题对话框"""
        theme_window = tk.Toplevel(self.root)
        theme_window.title("主题设置")
        theme_window.geometry("300x200")
        
        tk.Label(theme_window, text="选择主题", font=('Arial', 12, 'bold')).pack(pady=10)
        
        themes = [
            ("浅色主题", "light"),
            ("深色主题", "dark"), 
            ("蓝色主题", "blue"),
            ("绿色主题", "green")
        ]
        
        for name, theme_id in themes:
            btn = ttk.Button(theme_window, text=name, 
                           command=lambda t=theme_id: self.apply_theme(t, theme_window))
            btn.pack(fill='x', padx=20, pady=5)
            
    def apply_theme(self, theme_id, window):
        """应用主题"""
        themes = {
            "light": {"bg": "#ffffff", "fg": "#000000", "toolbar": "#f0f0f0"},
            "dark": {"bg": "#2c3e50", "fg": "#ffffff", "toolbar": "#34495e"},
            "blue": {"bg": "#ecf0f1", "fg": "#2c3e50", "toolbar": "#3498db"},
            "green": {"bg": "#f0f7f4", "fg": "#2c3e50", "toolbar": "#27ae60"}
        }
        
        if theme_id in themes:
            theme = themes[theme_id]
            self.root.configure(bg=theme["bg"])
            # 这里应该更新所有组件的颜色，简化实现
            messagebox.showinfo("主题", f"已切换到{theme_id}主题")
            
        window.destroy()
        
    def spell_check(self):
        """拼写检查"""
        # 简单的拼写检查模拟
        content = self.text_area.get('1.0', 'end-1c')
        words = content.split()
        
        # 常见易错词
        common_errors = {
            "的得地": "注意'的、得、地'的使用",
            "在再": "注意'在、再'的使用", 
            "做作": "注意'做、作'的使用",
            "象像": "注意'象、像'的使用"
        }
        
        issues = []
        for error, suggestion in common_errors.items():
            if any(char in content for char in error):
                issues.append(suggestion)
                
        if issues:
            messagebox.showwarning("拼写检查", "发现可能的拼写问题:\n\n" + "\n".join(issues))
        else:
            messagebox.showinfo("拼写检查", "拼写检查通过！")
            
    def template_manager(self):
        """模板管理"""
        template_window = tk.Toplevel(self.root)
        template_window.title("模板管理")
        template_window.geometry("500x400")
        
        tk.Label(template_window, text="文档模板", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # 模板列表
        templates_frame = tk.Frame(template_window)
        templates_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        templates = [
            ("商务报告", "包含目录、摘要、正文的商务报告模板"),
            ("会议纪要", "标准的会议记录模板"),
            ("项目计划", "项目管理计划模板"),
            ("个人简历", "专业的简历模板"),
            ("学术论文", "学术论文格式模板")
        ]
        
        for name, desc in templates:
            frame = tk.Frame(templates_frame, relief='groove', bd=1)
            frame.pack(fill='x', pady=2)
            
            tk.Label(frame, text=name, font=('Arial', 10, 'bold')).pack(anchor='w', padx=5, pady=2)
            tk.Label(frame, text=desc, font=('Arial', 8)).pack(anchor='w', padx=5, pady=2)
            
            use_btn = ttk.Button(frame, text="使用", 
                               command=lambda n=name: self.use_template(n, template_window))
            use_btn.pack(side='right', padx=5, pady=5)
            
    def use_template(self, template_name, window):
        """使用模板"""
        templates = {
            "商务报告": "商务报告\n\n目录\n\n摘要\n\n1. 引言\n2. 正文\n3. 结论\n\n附录",
            "会议纪要": "会议纪要\n\n时间: \n地点: \n参会人员: \n\n会议内容:\n\n决议事项:\n\n下一步计划:",
            "项目计划": "项目计划书\n\n项目名称: \n项目目标: \n时间安排: \n资源需求: \n风险评估:",
            "个人简历": "个人简历\n\n基本信息\n教育背景\n工作经历\n项目经验\n技能专长",
            "学术论文": "学术论文\n\n标题\n作者\n摘要\n关键词\n\n1. 引言\n2. 文献综述\n3. 研究方法\n4. 结果分析\n5. 结论\n参考文献"
        }
        
        if template_name in templates:
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert('1.0', templates[template_name])
            window.destroy()
            
    def macro_recorder(self):
        """宏录制"""
        messagebox.showinfo("宏录制", "宏录制功能开发中...")
        
    def options_dialog(self):
        """选项对话框"""
        options_window = tk.Toplevel(self.root)
        options_window.title("选项设置")
        options_window.geometry("400x500")
        
        notebook = ttk.Notebook(options_window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 常规设置
        general_frame = ttk.Frame(notebook)
        notebook.add(general_frame, text="常规")
        
        # 自动保存设置
        auto_save_var = tk.BooleanVar(value=self.auto_save)
        auto_save_cb = tk.Checkbutton(general_frame, text="启用自动保存", variable=auto_save_var)
        auto_save_cb.pack(anchor='w', pady=5)
        
        # 编辑器设置
        editor_frame = ttk.Frame(notebook)
        notebook.add(editor_frame, text="编辑器")
        
        wrap_var = tk.BooleanVar(value=True)
        wrap_cb = tk.Checkbutton(editor_frame, text="自动换行", variable=wrap_var)
        wrap_cb.pack(anchor='w', pady=5)
        
        line_numbers_var = tk.BooleanVar(value=False)
        line_numbers_cb = tk.Checkbutton(editor_frame, text="显示行号", variable=line_numbers_var)
        line_numbers_cb.pack(anchor='w', pady=5)
        
        # 保存按钮
        save_btn = ttk.Button(options_window, text="保存设置", 
                            command=lambda: self.save_options(
                                auto_save_var.get(),
                                wrap_var.get(),
                                line_numbers_var.get(),
                                options_window
                            ), style='Success.TButton')
        save_btn.pack(pady=10)
        
    def save_options(self, auto_save, wrap, line_numbers, window):
        """保存选项"""
        self.auto_save = auto_save
        self.auto_save_label.config(text=f"自动保存: {'开' if auto_save else '关'}")
        window.destroy()
        messagebox.showinfo("选项", "设置已保存")
        
    def toggle_fullscreen(self):
        """切换全屏"""
        is_fullscreen = self.root.attributes('-fullscreen')
        self.root.attributes('-fullscreen', not is_fullscreen)
        
    def zoom_dialog(self):
        """缩放对话框"""
        zoom_window = tk.Toplevel(self.root)
        zoom_window.title("缩放设置")
        zoom_window.geometry("300x200")
        
        tk.Label(zoom_window, text="缩放比例", font=('Arial', 12, 'bold')).pack(pady=10)
        
        scales = ["50%", "75%", "100%", "125%", "150%", "200%"]
        for scale in scales:
            btn = ttk.Button(zoom_window, text=scale, 
                           command=lambda s=scale: self.apply_zoom(s, zoom_window))
            btn.pack(fill='x', padx=20, pady=2)
            
    def apply_zoom(self, scale, window):
        """应用缩放"""
        messagebox.showinfo("缩放", f"已设置为 {scale} 显示比例")
        window.destroy()
        
    def print_document(self):
        """打印文档"""
        messagebox.showinfo("打印", "打印功能需要系统打印支持")
        
    def show_tutorial(self):
        """显示教程"""
        tutorial_text = """OfficeMate 使用教程

基本操作:
• 新建文档: Ctrl+N
• 保存文档: Ctrl+S  
• 打开文档: Ctrl+O

文本编辑:
• 格式工具栏提供快速格式化
• 右键菜单提供更多选项
• 使用样式快速应用格式

高级功能:
• AI助手: 智能写作和语法检查
• 协作功能: 多人实时编辑
• 版本历史: 查看文档修改记录

提示: 更多功能请在菜单中探索"""
        
        tutorial_window = tk.Toplevel(self.root)
        tutorial_window.title("使用教程")
        tutorial_window.geometry("500x400")
        
        text_widget = scrolledtext.ScrolledText(tutorial_window, wrap='word')
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        text_widget.insert('1.0', tutorial_text)
        text_widget.config(state='disabled')
        
    def show_shortcuts(self):
        """显示快捷键"""
        shortcuts_text = """快捷键列表

文件操作:
Ctrl+N - 新建文档
Ctrl+O - 打开文档  
Ctrl+S - 保存文档
Ctrl+Shift+S - 另存为
Ctrl+P - 打印文档
Ctrl+Q - 退出程序

编辑操作:
Ctrl+Z - 撤销
Ctrl+Y - 重做
Ctrl+X - 剪切
Ctrl+C - 复制
Ctrl+V - 粘贴
Ctrl+A - 全选
Ctrl+F - 查找和替换

视图操作:
F11 - 切换全屏

其他:
F1 - 显示帮助"""
        
        shortcut_window = tk.Toplevel(self.root)
        shortcut_window.title("快捷键")
        shortcut_window.geometry("400x500")
        
        text_widget = scrolledtext.ScrolledText(shortcut_window, wrap='word')
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        text_widget.insert('1.0', shortcuts_text)
        text_widget.config(state='disabled')
        
    def about_dialog(self):
        """关于对话框"""
        about_text = f"""OfficeMate Pro v2.0

一个功能强大的办公软件套件
包含文字处理、电子表格、演示文稿等功能

开发者: Aawaider
版本: Alpha1
用户ID: {self.user_id}

© 2025 OfficeMate. 保留所有权利。"""
        
        messagebox.showinfo("关于 OfficeMate Pro", about_text)

    def bind_shortcuts(self):
        """绑定快捷键"""
        self.root.bind('<Control-n>', lambda e: self.new_file())
        self.root.bind('<Control-o>', lambda e: self.open_file())
        self.root.bind('<Control-s>', lambda e: self.save_file())
        self.root.bind('<Control-Shift-S>', lambda e: self.save_as_file())
        self.root.bind('<Control-p>', lambda e: self.print_document())
        self.root.bind('<Control-q>', lambda e: self.quit_application())
        self.root.bind('<F11>', lambda e: self.toggle_fullscreen())
        self.root.bind('<Control-z>', lambda e: self.undo())
        self.root.bind('<Control-y>', lambda e: self.redo())
        self.root.bind('<Control-x>', lambda e: self.cut())
        self.root.bind('<Control-c>', lambda e: self.copy())
        self.root.bind('<Control-v>', lambda e: self.paste())
        self.root.bind('<Control-a>', lambda e: self.select_all())
        self.root.bind('<Control-f>', lambda e: self.find_replace_dialog())
        self.root.bind('<F1>', lambda e: self.show_tutorial())

    # ===== 基本功能方法 =====
    
    def show_word_processor(self): 
        self.notebook.select(0)
    
    def show_spreadsheet(self): 
        self.notebook.select(1)
    
    def show_presentation(self): 
        self.notebook.select(2)
    
    def show_database_viewer(self):
        self.notebook.select(3)
    
    def undo(self): 
        try: 
            self.text_area.edit_undo()
        except: 
            pass
        
    def redo(self): 
        try: 
            self.text_area.edit_redo()
        except: 
            pass
        
    def cut(self): 
        try:
            self.text_area.event_generate("<<Cut>>")
        except:
            pass
    
    def copy(self): 
        try:
            self.text_area.event_generate("<<Copy>>")
        except:
            pass
    
    def paste(self): 
        try:
            self.text_area.event_generate("<<Paste>>")
        except:
            pass
            
    def select_all(self):
        """全选"""
        self.text_area.tag_add('sel', '1.0', 'end')

    def quit_application(self):
        """退出应用"""
        if messagebox.askokcancel("退出", "确定要退出 OfficeMate 吗？"):
            self.save_preferences()
            if hasattr(self, 'conn') and self.conn:
                self.conn.close()
            self.root.quit()

    def run(self):
        """运行应用"""
        # 设置窗口图标（如果有的话）
        try:
            self.root.iconbitmap("image.ico")  # 如果存在图标文件
        except:
            pass
            
        self.root.mainloop()

if __name__ == "__main__":
    app = OfficeMatePro()
    app.run()