#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel导入MySQL工具主应用
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import sys
import os
import requests
import json

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app.excel_reader import ExcelReader
from app.db_handler import DBHandler
from app.import_manager import ImportManager

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('excel_to_mysql.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ExcelToMysqlApp(tk.Tk):
    """Excel导入MySQL应用主类"""
    
    def __init__(self):
        """初始化应用"""
        super().__init__()
        
        # 应用配置
        self.title("Excel导入MySQL工具")
        self.geometry("1000x700")
        self.minsize(800, 600)
        
        # 全局变量
        self.excel_file = None
        self.excel_reader = None
        self.db_handler = None
        self.import_manager = None
        self.current_table = None
        
        # 初始化界面
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        # 创建主画布用于实现整体页面滚动
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 放置画布和滚动条
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # 创建主框架，放置在画布内部
        main_frame = ttk.Frame(self.canvas, padding="10")
        
        # 将主框架添加到画布
        self.canvas.create_window((0, 0), window=main_frame, anchor="nw", tags="main_frame")
        
        # 配置网格权重
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(9, weight=1)  # 字段映射框架
        main_frame.rowconfigure(12, weight=1)  # 日志框架
        
        # 监听主框架大小变化，更新画布滚动区域
        def on_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        main_frame.bind("<Configure>", on_configure)
        
        # 绑定鼠标滚轮事件，支持上下滚动
        def _on_mousewheel(event):
            # Windows平台使用event.delta，Mac使用event.num
            delta = event.delta // 120 if event.delta else -1 if event.num == 5 else 1
            self.canvas.yview_scroll(-delta, "units")  # 反转滚动方向，使鼠标滚轮体验更自然
        
        # 绑定鼠标滚轮事件到画布，而不是整个应用
        # 这样当鼠标在下拉框上时，滚轮不会影响下拉框选择
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # 为Linux/Mac平台添加滚轮支持（Button-4是上滚，Button-5是下滚）
        def _on_button_wheel(event):
            if event.num == 4:  # 上滚
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:  # 下滚
                self.canvas.yview_scroll(1, "units")
        
        self.canvas.bind("<Button-4>", _on_button_wheel)
        self.canvas.bind("<Button-5>", _on_button_wheel)
        
        # 数据库连接配置
        ttk.Label(main_frame, text="数据库连接配置", font=("Arial", 12, "bold")).grid(
            row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 10)
        )
        
        # 主机名
        ttk.Label(main_frame, text="主机名:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)
        self.host_entry = ttk.Entry(main_frame)
        self.host_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.host_entry.insert(0, "localhost")
        
        # 端口
        ttk.Label(main_frame, text="端口:").grid(row=1, column=2, sticky=tk.E, padx=5, pady=5)
        self.port_entry = ttk.Entry(main_frame, width=10)
        self.port_entry.grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)
        self.port_entry.insert(0, "3306")
        
        # 用户名
        ttk.Label(main_frame, text="用户名:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=5)
        self.user_entry = ttk.Entry(main_frame)
        self.user_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.user_entry.insert(0, "root")
        
        # 密码
        ttk.Label(main_frame, text="密码:").grid(row=2, column=2, sticky=tk.E, padx=5, pady=5)
        self.password_entry = ttk.Entry(main_frame, show="*")
        self.password_entry.grid(row=2, column=3, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 数据库名
        ttk.Label(main_frame, text="数据库:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=5)
        self.db_entry = ttk.Entry(main_frame)
        self.db_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 连接按钮
        self.connect_btn = ttk.Button(main_frame, text="连接数据库", command=self.connect_db)
        self.connect_btn.grid(row=3, column=2, columnspan=2, pady=5, padx=5)
        
        # Excel文件选择
        ttk.Label(main_frame, text="Excel文件选择", font=("Arial", 12, "bold")).grid(
            row=4, column=0, columnspan=4, sticky=tk.W, pady=(15, 10)
        )
        
        ttk.Label(main_frame, text="选择文件:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=5)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.excel_path_var, state="readonly").grid(
            row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Button(main_frame, text="浏览", command=self.browse_excel).grid(
            row=5, column=3, padx=5, pady=5
        )
        
        # sheet选择
        ttk.Label(main_frame, text="Sheet:").grid(row=6, column=0, sticky=tk.E, padx=5, pady=5)
        self.sheet_combo = ttk.Combobox(main_frame, state="readonly")
        self.sheet_combo.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # 表选择
        ttk.Label(main_frame, text="目标表:").grid(row=6, column=2, sticky=tk.E, padx=5, pady=5)
        self.table_combo = ttk.Combobox(main_frame, state="readonly")
        self.table_combo.grid(row=6, column=3, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.table_combo.bind("<<ComboboxSelected>>", self.on_table_selected)
        
        # DeepSeek API配置
        api_frame = ttk.LabelFrame(main_frame, text="DeepSeek API配置", padding="10")
        api_frame.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        # 配置列权重，与主框架保持一致
        api_frame.columnconfigure(1, weight=1)
        
        ttk.Label(api_frame, text="API密钥:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        self.deepseek_api_key = ttk.Entry(api_frame, show="*")
        self.deepseek_api_key.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.deepseek_api_key.insert(0, "")  # 默认空值
        
        # 自动匹配按钮 - 放在输入框后面
        self.auto_match_btn = ttk.Button(
            api_frame, 
            text="自动匹配字段", 
            command=self.auto_match_fields
        )
        self.auto_match_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 字段映射
        ttk.Label(main_frame, text="字段映射", font=("Arial", 12, "bold")).grid(
            row=8, column=0, columnspan=4, sticky=tk.W, pady=(15, 10)
        )
        
        # 映射框架
        mapping_frame = ttk.LabelFrame(main_frame, text="字段映射", padding="10")
        mapping_frame.grid(
            row=9, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5
        )
        mapping_frame.columnconfigure(0, weight=1)
        mapping_frame.columnconfigure(1, weight=1)
        mapping_frame.rowconfigure(0, weight=0)  # 标题行不扩展
        mapping_frame.rowconfigure(1, weight=0)  # 状态行不扩展
        mapping_frame.rowconfigure(2, weight=1)  # 映射内容行扩展
        
        # 保存引用
        self.mapping_frame = mapping_frame
        
        # 初始化映射组件列表
        self.mapping_widgets = []
        
        # 映射标题
        ttk.Label(mapping_frame, text="Excel列", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(mapping_frame, text="数据库字段", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)
        
        # 唯一索引设置
        settings_frame = ttk.Frame(mapping_frame)
        settings_frame.grid(row=0, column=3, sticky=(tk.N, tk.W), padx=10, pady=5)
        
        ttk.Label(settings_frame, text="唯一索引:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.unique_combo = ttk.Combobox(settings_frame, state="readonly", width=20)
        self.unique_combo.grid(row=0, column=1, padx=5, pady=5)
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=10, column=0, columnspan=4, pady=15)
        button_frame.columnconfigure(0, weight=1)
        
        # 导入按钮
        self.import_btn = ttk.Button(
            button_frame, 
            text="开始导入", 
            command=self.start_import,
            style="Accent.TButton"
        )
        self.import_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # 日志输出
        ttk.Label(main_frame, text="导入日志", font=(("Arial", 12, "bold"))).grid(
            row=11, column=0, columnspan=4, sticky=tk.W, pady=(15, 10)
        )
        
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="10")
        log_frame.grid(
            row=12, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5
        )
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 日志文本框
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 日志滚动条
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), padx=5, pady=5)
        
        # 样式配置
        self.style = ttk.Style()
        self.style.configure("Accent.TButton", foreground="red")
    
    def log(self, message):
        """记录日志
        
        Args:
            message (str): 日志消息
        """
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        logger.info(message)
    
    def connect_db(self):
        """连接数据库"""
        try:
            # 获取连接信息
            host = self.host_entry.get()
            port = int(self.port_entry.get())
            user = self.user_entry.get()
            password = self.password_entry.get()
            database = self.db_entry.get()
            
            # 初始化数据库处理器
            self.db_handler = DBHandler(host, port, user, password, database)
            
            # 连接数据库
            self.db_handler.connect()
            self.log("成功连接到数据库")
            
            # 加载表
            tables = self.db_handler.get_tables()
            self.table_combo['values'] = tables
            self.log(f"成功加载{len(tables)}个表")
            
        except Exception as e:
            messagebox.showerror("错误", f"数据库连接失败: {str(e)}")
            self.log(f"数据库连接失败: {str(e)}")
    
    def browse_excel(self):
        """浏览选择Excel文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls *.xlsm")],
            title="选择Excel文件"
        )
        
        if file_path:
            self.excel_file = file_path
            self.excel_path_var.set(file_path)
            self.log(f"选择了Excel文件: {file_path}")
            
            # 初始化Excel读取器
            try:
                self.excel_reader = ExcelReader(file_path)
                sheet_names = self.excel_reader.load_file()
                self.sheet_combo['values'] = sheet_names
                
                # 默认选择第一个sheet
                if sheet_names:
                    self.sheet_combo.set(sheet_names[0])
                    self.on_sheet_selected(None)
                    
            except Exception as e:
                messagebox.showerror("错误", f"加载Excel文件失败: {str(e)}")
                self.log(f"加载Excel文件失败: {str(e)}")
    
    def on_sheet_selected(self, event):
        """当选择sheet时的处理"""
        if not self.excel_reader:
            return
        
        sheet_name = self.sheet_combo.get()
        if not sheet_name:
            return
        
        try:
            # 读取sheet数据
            self.excel_reader.read_sheet(sheet_name)
            self.log(f"成功读取sheet: {sheet_name}")
            
            # 更新字段映射
            self.update_field_mapping()
            
        except Exception as e:
            messagebox.showerror("错误", f"读取sheet失败: {str(e)}")
            self.log(f"读取sheet失败: {str(e)}")
    
    def on_table_selected(self, event):
        """当选择表时的处理"""
        if not self.db_handler:
            return
        
        table_name = self.table_combo.get()
        if not table_name:
            return
        
        try:
            self.current_table = table_name
            self.log(f"选择了表: {table_name}")
            
            # 更新字段映射
            self.update_field_mapping()
            
            # 更新唯一索引选项
            columns = self.db_handler.get_table_columns(table_name)
            self.unique_combo['values'] = [''] + columns
            
        except Exception as e:
            messagebox.showerror("错误", f"加载表字段失败: {str(e)}")
            self.log(f"加载表字段失败: {str(e)}")
    
    def update_field_mapping(self):
        """更新字段映射表"""
        self.log("开始更新字段映射")
        
        # 清空现有映射组件（从第1行开始，保留标题行）
        for widget in self.mapping_frame.winfo_children():
            grid_info = widget.grid_info()
            if grid_info['row'] > 0:
                widget.destroy()
        
        # 重置映射组件列表
        self.mapping_widgets = []
        
        # 显示当前状态
        status_text = "字段映射状态: "
        
        # 检查各变量状态
        excel_cols = []
        db_cols = []
        
        if not self.excel_reader:
            status_text += "请先选择Excel文件"
        else:
            try:
                excel_cols = self.excel_reader.get_columns()
                status_text += f"已选择Excel文件，包含{len(excel_cols)}列"
            except Exception as e:
                status_text += f"获取Excel列失败: {str(e)}"
                self.log(f"获取Excel列失败: {str(e)}")
        
        if not self.db_handler:
            status_text += " | 请先连接数据库"
        else:
            status_text += " | 已连接数据库"
        
        if not self.current_table:
            status_text += " | 请先选择目标表"
        else:
            try:
                db_cols = self.db_handler.get_table_columns(self.current_table)
                status_text += f" | 已选择表: {self.current_table}，包含{len(db_cols)}个字段"
            except Exception as e:
                status_text += f" | 获取表字段失败: {str(e)}"
                self.log(f"获取表字段失败: {str(e)}")
        
        # 显示状态
        status_label = ttk.Label(self.mapping_frame, text=status_text, foreground="blue")
        status_label.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)
        
        # 简化条件判断，只要有Excel列和数据库字段就显示映射
        if excel_cols and db_cols:
            self.log(f"显示字段映射: Excel列数={len(excel_cols)}, 数据库字段数={len(db_cols)}")
            
            # 添加映射行
            for i, excel_col in enumerate(excel_cols, start=2):
                # 创建Excel列标签
                excel_label = ttk.Label(self.mapping_frame, text=excel_col, anchor=tk.W)
                excel_label.grid(row=i, column=0, padx=10, pady=5, sticky=tk.W)
                
                # 创建数据库字段下拉框
                db_combo = ttk.Combobox(self.mapping_frame, state="readonly", width=30)
                db_combo['values'] = [''] + db_cols
                
                # 自动匹配相同名称的字段
                matched_col = excel_col if excel_col in db_cols else ''
                db_combo.set(matched_col)
                
                db_combo.grid(row=i, column=1, padx=10, pady=5, sticky=tk.W)
                
                # 存储映射关系
                self.mapping_widgets.append((excel_col, db_combo))
        else:
            # 显示提示信息
            hint_text = "请完成以下步骤:"
            hint_label = ttk.Label(self.mapping_frame, text=hint_text, foreground="green", font=(("Arial", 10, "bold")))
            hint_label.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
            
            if not self.excel_reader:
                step_label = ttk.Label(self.mapping_frame, text="1. 点击'浏览'按钮选择Excel文件", foreground="black")
                step_label.grid(row=3, column=0, padx=20, pady=2, sticky=tk.W)
            
            if not self.db_handler:
                step_label = ttk.Label(self.mapping_frame, text="2. 填写数据库连接信息并点击'连接'按钮", foreground="black")
                step_label.grid(row=4, column=0, padx=20, pady=2, sticky=tk.W)
            
            if not self.current_table:
                step_label = ttk.Label(self.mapping_frame, text="3. 从'表名'下拉框中选择目标表", foreground="black")
                step_label.grid(row=5, column=0, padx=20, pady=2, sticky=tk.W)
            
            if self.excel_reader and not excel_cols:
                step_label = ttk.Label(self.mapping_frame, text="4. 从'Sheet'下拉框中选择要导入的工作表", foreground="black")
                step_label.grid(row=6, column=0, padx=20, pady=2, sticky=tk.W)
        
        self.log("字段映射更新完成")
    
    def get_field_mapping(self):
        """获取字段映射
        
        Returns:
            list: 字段映射关系列表，每个元素为 (excel_col, db_col) 元组
        """
        mapping = []
        for excel_col, db_combo in self.mapping_widgets:
            db_col = db_combo.get()
            if db_col:
                mapping.append((excel_col, db_col))
        return mapping
    
    def auto_match_fields(self):
        """调用DeepSeek API自动匹配Excel和数据库字段"""
        try:
            # 检查必要条件
            if not self.excel_reader:
                messagebox.showwarning("警告", "请先选择Excel文件")
                return
            
            if not self.db_handler:
                messagebox.showwarning("警告", "请先连接数据库")
                return
            
            if not self.current_table:
                messagebox.showwarning("警告", "请先选择目标表")
                return
            
            # 获取API密钥
            api_key = self.deepseek_api_key.get()
            if not api_key:
                messagebox.showwarning("警告", "请先配置DeepSeek API密钥")
                return
            
            # 获取Excel列和数据库字段
            excel_cols = self.excel_reader.get_columns()
            db_cols = self.db_handler.get_table_columns(self.current_table)
            
            # 构建请求数据
            prompt = f"请帮我匹配以下Excel列和数据库字段，返回格式为JSON，键是Excel列名，值是匹配的数据库字段名。\n\nExcel列: {excel_cols}\n\n数据库字段: {db_cols}\n\n要求:\n1. 只返回JSON格式，不要其他任何内容\n2. 对于不匹配的字段，值设为空字符串\n3. 尽可能根据语义匹配，而不仅仅是名称匹配"
            
            # 调用DeepSeek API
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "gzip, deflate, br"
            }
            data = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": "你是一个专业的数据分析师，擅长匹配Excel列和数据库字段。"},
                    {"role": "user", "content": prompt}
                ],
                "max_tokens": 1024,
                "temperature": 0.1
            }
            
            self.log("正在调用DeepSeek API进行字段匹配...")
            try:
                response = requests.post(
                    "https://api.deepseek.com/v1/chat/completions",
                    headers=headers,
                    json=data,
                    timeout=30
                )
                response.raise_for_status()
                response.encoding = 'utf-8'
            except UnicodeEncodeError as e:
                # 处理非ASCII字符编码错误
                self.log(f"API调用失败: 编码错误 - {str(e)}")
                messagebox.showerror(
                    "错误", 
                    "自动匹配字段失败: API密钥中包含非法字符\n\nHTTP请求头只能包含ASCII字符，请确保API密钥格式正确。"
                )
                return
            except requests.exceptions.HTTPError as e:
                # 重新抛出HTTP错误，由上层处理
                raise
            
            # 解析响应
            result = response.json()
            self.log(f"API原始响应: {result}")
            
            try:
                # 获取AI回复内容
                ai_content = result["choices"][0]["message"]["content"]
                self.log(f"AI回复内容: {ai_content}")
                
                # 处理Markdown代码块，移除 ```json 和 ``` 标记
                if ai_content.startswith('```json') and ai_content.endswith('```'):
                    # 提取JSON内容，去除首尾的代码块标记
                    json_content = ai_content[7:-3].strip()
                    self.log(f"提取的JSON内容: {json_content}")
                    # 解析JSON
                    matching = json.loads(json_content)
                else:
                    # 直接解析JSON
                    matching = json.loads(ai_content)
                self.log(f"字段匹配结果: {matching}")
                
                # 应用匹配结果
                for i, (excel_col, db_combo) in enumerate(self.mapping_widgets):
                    if excel_col in matching and matching[excel_col] in db_cols:
                        db_combo.set(matching[excel_col])
                
                messagebox.showinfo("成功", "字段自动匹配完成！你可以手动调整不匹配的字段。")
                self.log("字段自动匹配完成")
            except json.JSONDecodeError as e:
                # 处理JSON解析错误
                self.log(f"JSON解析失败: {str(e)}")
                self.log(f"AI回复内容: {ai_content}")
                messagebox.showerror("错误", f"自动匹配字段失败: AI返回内容不是有效的JSON格式\n{ai_content[:200]}...")
            except KeyError as e:
                # 处理API响应格式错误
                self.log(f"API响应格式错误: {str(e)}")
                messagebox.showerror("错误", f"自动匹配字段失败: API响应格式错误\n{str(e)}")
            
        except requests.exceptions.HTTPError as e:
            # 处理HTTP错误
            if e.response.status_code == 401:
                # 处理API密钥无效的情况
                self.log(f"API调用失败: 无效的API密钥\n{str(e)}")
                messagebox.showerror(
                    "错误", 
                    "自动匹配字段失败: API密钥无效或已过期\n\n请检查您的DeepSeek API密钥是否正确配置。\n可以通过DeepSeek官网获取有效的API密钥。"
                )
            else:
                # 其他HTTP错误
                self.log(f"API调用失败: HTTP错误 {e.response.status_code}\n{str(e)}")
                try:
                    # 尝试获取API返回的错误详情
                    error_detail = e.response.json()
                    self.log(f"API错误详情: {error_detail}")
                    messagebox.showerror(
                        "错误", 
                        f"自动匹配字段失败: HTTP错误 {e.response.status_code}\n\n详情: {error_detail.get('error', {}).get('message', '未知错误')}"
                    )
                except:
                    # 如果无法解析错误详情，显示基本错误信息
                    messagebox.showerror(
                        "错误", 
                        f"自动匹配字段失败: HTTP错误 {e.response.status_code}\n\n{str(e)}"
                    )
        except requests.exceptions.RequestException as e:
            # 处理其他网络请求错误
            self.log(f"网络请求失败: {str(e)}")
            messagebox.showerror(
                "错误", 
                f"自动匹配字段失败: 网络请求错误\n\n请检查您的网络连接，或稍后重试。\n{str(e)}"
            )
        except Exception as e:
            # 处理其他错误
            self.log(f"自动匹配字段失败: {str(e)}")
            import traceback
            self.log(f"错误堆栈: {traceback.format_exc()}")
            messagebox.showerror("错误", f"自动匹配字段失败: {str(e)}")
    
    def start_import(self):
        """开始导入数据"""
        # 验证
        if not self.excel_reader:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        if not self.db_handler:
            messagebox.showwarning("警告", "请先连接数据库")
            return
        
        if not self.current_table:
            messagebox.showwarning("警告", "请选择目标表")
            return
        
        # 获取字段映射
        mapping = self.get_field_mapping()
        if not mapping:
            messagebox.showwarning("警告", "请设置字段映射")
            return
        
        # 获取唯一索引
        unique_key = self.unique_combo.get()
        
        try:
            # 初始化导入管理器
            self.import_manager = ImportManager(self.excel_reader, self.db_handler)
            
            # 设置导入参数
            self.import_manager.set_mapping(mapping)
            if unique_key:
                self.import_manager.set_unique_key(unique_key)
            
            # 开始导入
            self.log("开始导入数据...")
            result = self.import_manager.import_data(self.current_table)
            
            if result['success']:
                messagebox.showinfo("成功", f"数据导入成功，共导入{result['affected_rows']}行")
                self.log(f"数据导入成功，共导入{result['affected_rows']}行")
            else:
                messagebox.showerror("错误", result['message'])
                self.log(result['message'])
                
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {str(e)}")
            self.log(f"导入失败: {str(e)}")
    
    def on_closing(self):
        """关闭应用时的处理"""
        # 关闭数据库连接
        if self.db_handler:
            self.db_handler.close()
        self.destroy()

def main():
    """主函数"""
    app = ExcelToMysqlApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()

if __name__ == "__main__":
    main()
