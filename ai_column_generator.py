import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openai import OpenAI
import configparser
import json
from typing import List, Dict, Any, Optional
import time
import requests  # 用于Ollama API请求
import threading  # 用于多线程处理
import queue  # 用于线程间通信
import re
# chardet会在需要时动态导入

class TemplatePreviewDialog:
    """模板预览对话框，用于预览模板效果和测试变量替换"""
    def __init__(self, parent, template_text, variables=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("模板预览")
        self.dialog.geometry("600x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.template_text = template_text
        self.variables = variables or {}
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 变量区域
        var_frame = ttk.LabelFrame(main_frame, text="变量测试", padding=10)
        var_frame.pack(fill=tk.X, pady=5)
        
        # 添加常用变量输入框
        self.var_entries = {}
        row = 0
        
        # 引用内容变量
        ttk.Label(var_frame, text="引用内容:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ref_content_frame = ttk.Frame(var_frame)
        ref_content_frame.grid(row=row, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        self.ref_content_var = tk.StringVar(value=self.variables.get("引用内容", "这是测试引用内容"))
        ref_content_entry = ttk.Entry(ref_content_frame, textvariable=self.ref_content_var, width=40)
        ref_content_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 模板区域
        template_frame = ttk.LabelFrame(main_frame, text="模板内容", padding=10)
        template_frame.pack(fill=tk.X, pady=5)
        
        template_scroll = ttk.Scrollbar(template_frame)
        template_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.template_text_widget = tk.Text(template_frame, height=6, wrap=tk.WORD, yscrollcommand=template_scroll.set)
        self.template_text_widget.pack(fill=tk.X, expand=True)
        self.template_text_widget.insert(tk.END, self.template_text)
        template_scroll.config(command=self.template_text_widget.yview)
        
        # 添加更新按钮
        update_frame = ttk.Frame(main_frame)
        update_frame.pack(fill=tk.X, pady=5)
        ttk.Button(update_frame, text="更新预览", command=self.update_preview).pack(side=tk.LEFT, padx=5)
        ttk.Label(update_frame, text="(修改模板或变量后点击更新)").pack(side=tk.LEFT, padx=5)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="预览结果", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        preview_scroll = ttk.Scrollbar(preview_frame)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD, yscrollcommand=preview_scroll.set)
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        preview_scroll.config(command=self.preview_text.yview)
        
        # 底部按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="应用到编辑器", command=self.apply_to_editor).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=self.dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 初始预览
        self.update_preview()
        
    def update_preview(self):
        """更新预览内容"""
        # 获取当前模板内容（允许在预览窗口中编辑）
        current_template = self.template_text_widget.get("1.0", tk.END).strip()
        
        # 收集变量
        variables = {
            "引用内容": self.ref_content_var.get().strip(),
        }
        
        # 替换变量生成预览
        try:
            # 假设父窗口是AIColumnGenerator实例
            parent = self.dialog.master
            if hasattr(parent, 'replace_template_variables'):
                preview = parent.replace_template_variables(current_template, variables)
            else:
                # 简单替换逻辑(备用)
                preview = current_template
                for var_name, var_value in variables.items():
                    preview = preview.replace(f"{{{var_name}}}", var_value)
                    
                    # 处理条件变量
                    if var_value:
                        pattern = r'\{如果:' + re.escape(var_name) + r':(.*?)\}'
                        preview = re.sub(pattern, r'\1', preview)
                    else:
                        pattern = r'\{如果:' + re.escape(var_name) + r':.*?\}'
                        preview = re.sub(pattern, '', preview)
            
            # 更新预览文本
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, preview)
            
        except Exception as e:
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, f"预览生成错误: {str(e)}")
    
    def apply_to_editor(self):
        """将当前编辑的模板应用到主编辑器"""
        # 获取当前模板内容
        current_template = self.template_text_widget.get("1.0", tk.END).strip()
        
        # 确认应用
        if not messagebox.askyesno("确认应用", 
                               "确定要将当前编辑的模板应用到主编辑器吗？\n这将覆盖现有的提示词内容。"):
            return
            
        # 应用到主窗口的提示词编辑器
        parent = self.dialog.master
        if hasattr(parent, 'prompt_text'):
            parent.prompt_text.delete("1.0", tk.END)
            parent.prompt_text.insert(tk.END, current_template)
            messagebox.showinfo("成功", "模板已应用到编辑器")
            self.dialog.destroy()
        else:
            messagebox.showerror("错误", "无法应用模板到编辑器")

class AIColumnGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("LocalAItable")
        self.root.geometry("900x750")  # 增加窗口高度
        
        self.file_path = None
        self.df = None
        self.api_key = None
        self.ollama_url = "http://localhost:11434"  # Ollama默认API地址
        self.api_type = "openai"  # 默认使用OpenAI API
        
        # 提示词模板管理
        self.templates = {}  # 用户保存的模板
        self.preset_templates = {
            "摘要生成": "请根据以下内容生成一段简洁的摘要：\n\n{引用内容}",
            "内容翻译": "请将以下内容翻译成英文：\n\n{引用内容}",
            "情感分析": "请分析以下内容的情感倾向(积极、消极或中性)，并给出理由：\n\n{引用内容}",
            "关键词提取": "请从以下内容中提取5个最重要的关键词或短语：\n\n{引用内容}",
            "内容分类": "请将以下内容分类到最合适的类别(例如：科技、健康、教育、娱乐等)，并说明理由：\n\n{引用内容}",
            "观点提取": "请从以下内容中提取主要观点和论点：\n\n{引用内容}",
            "问题回答": "请根据以下内容回答问题：\n{如果:问题:问题: {问题}\n\n}参考内容:\n{引用内容}",
            "数据提取": "请从以下文本中提取所有数字数据，并按类别整理：\n\n{引用内容}",
            "简介生成": "请根据以下内容生成一段专业的产品/服务简介，突出其主要特点和价值：\n\n{引用内容}",
            "医学报告分析": "请分析以下医学报告内容，提取关键指标并解释其含义：\n\n{引用内容}",
            "技术文档简化": "请将以下技术文档内容转换为普通用户容易理解的语言：\n\n{引用内容}",
            "学术摘要": "请将以下学术内容生成一段摘要，包含研究目的、方法、结果和结论：\n\n{引用内容}",
            "血压数据提取": "请从以下医疗记录中提取患者的血压数据，只返回血压值，格式为'血压X/YmmHg'：\n\n{引用内容}"
        }
        
        self.load_config()
        self.load_templates()
        
        self.create_widgets()
        
        # 窗口关闭时的处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        """窗口关闭时的处理"""
        # 解绑所有事件，避免内存泄漏
        if hasattr(self, 'main_canvas'):
            self.main_canvas.unbind_all("<MouseWheel>")
            self.main_canvas.unbind_all("<Button-4>")
            self.main_canvas.unbind_all("<Button-5>")
        
        # 关闭窗口
        self.root.destroy()
    
    def load_config(self):
        """加载配置文件，获取API密钥和Ollama设置"""
        config = configparser.ConfigParser()
        
        if os.path.exists('config.ini'):
            config.read('config.ini')
            if 'API' in config:
                if 'openai_api_key' in config['API']:
                    self.api_key = config['API']['openai_api_key']
                if 'ollama_url' in config['API']:
                    self.ollama_url = config['API']['ollama_url']
                if 'api_type' in config['API']:
                    self.api_type = config['API']['api_type']
        
        if not self.api_key and self.api_type == 'openai':
            # 如果没有找到API密钥，则从环境变量中获取
            self.api_key = os.environ.get('OPENAI_API_KEY')
    
    def save_config(self):
        """保存配置到文件"""
        config = configparser.ConfigParser()
        config['API'] = {
            'openai_api_key': self.api_key if self.api_key else '',
            'ollama_url': self.ollama_url,
            'api_type': self.api_type
        }
        
        with open('config.ini', 'w') as f:
            config.write(f)
    
    def load_templates(self):
        """加载用户保存的提示词模板"""
        if os.path.exists('prompt_templates.json'):
            try:
                with open('prompt_templates.json', 'r', encoding='utf-8') as f:
                    self.templates = json.load(f)
            except:
                self.templates = {}
    
    def save_templates(self):
        """保存提示词模板到文件"""
        with open('prompt_templates.json', 'w', encoding='utf-8') as f:
            json.dump(self.templates, f, ensure_ascii=False, indent=2)
    
    def create_widgets(self):
        """创建GUI组件"""
        # 创建状态栏（放在最底部且不会滚动）
        self.status_var = tk.StringVar(value="本项目由华中科技大学同济医院影像科刘远康博士独立开发，免费开源。如有定制化或其他商业需求请加微信Blackbirdflyinthesky")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建canvas框架来容纳滚动区域
        canvas_frame = ttk.Frame(self.root)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # 创建一个主滚动框架包含所有内容
        self.main_canvas = tk.Canvas(canvas_frame)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加垂直滚动条
        self.main_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 配置Canvas
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        
        # 创建内容框架
        self.scrollable_frame = ttk.Frame(self.main_canvas)
        canvas_window = self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # 当框架大小变化时更新Canvas滚动区域
        def _on_frame_configure(event):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        
        self.scrollable_frame.bind("<Configure>", _on_frame_configure)
        
        # 当Canvas大小改变时，更新窗口宽度以匹配Canvas宽度
        def _on_canvas_configure(event):
            self.main_canvas.itemconfig(canvas_window, width=event.width)
        
        self.main_canvas.bind("<Configure>", _on_canvas_configure)
        
        # 添加鼠标滚轮滚动支持
        def _on_mousewheel(event):
            # 处理不同平台的鼠标滚轮事件
            if hasattr(event, 'delta'):  # Windows
                delta = -1 * (event.delta // 120)
            elif hasattr(event, 'num'):  # Linux
                if event.num == 4:
                    delta = -1
                elif event.num == 5:
                    delta = 1
                else:
                    return
            else:
                return
                
            self.main_canvas.yview_scroll(delta, "units")
        
        # Windows滚轮绑定
        self.main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # Linux滚轮绑定
        self.main_canvas.bind_all("<Button-4>", _on_mousewheel)
        self.main_canvas.bind_all("<Button-5>", _on_mousewheel)
        
        # 创建主框架 - 所有组件都放在这个框架里
        main_frame = ttk.Frame(self.scrollable_frame, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择部分
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=10)
        file_frame.pack(fill=tk.X, pady=5)
        
        file_button_frame = ttk.Frame(file_frame)
        file_button_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        ttk.Button(file_button_frame, text="选择Excel/CSV文件", command=self.select_file).pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(file_button_frame, text="手动指定编码打开", command=self.select_file_with_encoding).pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(file_button_frame, text="编码问题帮助", command=self.show_encoding_help).pack(fill=tk.X, padx=5, pady=2)
        
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # API配置部分
        api_frame = ttk.LabelFrame(main_frame, text="API配置", padding=10)
        api_frame.pack(fill=tk.X, pady=5)
        
        # API类型选择
        ttk.Label(api_frame, text="AI服务类型:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.api_type_var = tk.StringVar(value=self.api_type)
        api_type_combo = ttk.Combobox(api_frame, textvariable=self.api_type_var, width=20, state="readonly")
        api_type_combo['values'] = ["openai", "ollama"]
        api_type_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        api_type_combo.bind("<<ComboboxSelected>>", self.on_api_type_changed)
        
        # OpenAI API密钥
        self.openai_key_label = ttk.Label(api_frame, text="OpenAI API密钥:")
        self.openai_key_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.api_key_var = tk.StringVar(value=self.api_key if self.api_key else "")
        self.api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=50, show="*")
        self.api_key_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        self.save_api_btn = ttk.Button(api_frame, text="保存API密钥", command=self.save_api_key)
        self.save_api_btn.grid(row=1, column=2, padx=5, pady=5)
        self.toggle_api_btn = ttk.Button(api_frame, text="显示/隐藏", command=self.toggle_api_key_visibility)
        self.toggle_api_btn.grid(row=1, column=3, padx=5, pady=5)
        
        # Ollama URL
        self.ollama_url_label = ttk.Label(api_frame, text="Ollama URL:")
        self.ollama_url_label.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.ollama_url_var = tk.StringVar(value=self.ollama_url)
        self.ollama_url_entry = ttk.Entry(api_frame, textvariable=self.ollama_url_var, width=50)
        self.ollama_url_entry.grid(row=2, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        self.test_ollama_btn = ttk.Button(api_frame, text="测试连接", command=self.test_ollama_connection)
        self.test_ollama_btn.grid(row=2, column=2, padx=5, pady=5)
        self.list_ollama_btn = ttk.Button(api_frame, text="获取模型列表", command=self.list_ollama_models)
        self.list_ollama_btn.grid(row=2, column=3, padx=5, pady=5)
        
        # 模型选择
        ttk.Label(api_frame, text="选择模型:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.model_var = tk.StringVar(value="gpt-3.5-turbo")
        self.model_combo = ttk.Combobox(api_frame, textvariable=self.model_var, width=40)
        self.model_combo['values'] = ["gpt-3.5-turbo", "gpt-4", "gpt-4-turbo"]
        self.model_combo.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 根据初始API类型显示/隐藏相关控件
        self.update_api_widgets()
        
        # 列选择部分
        columns_frame = ttk.LabelFrame(main_frame, text="列选择", padding=10)
        columns_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(columns_frame, text="写入列:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_column_var = tk.StringVar()
        self.target_column_combo = ttk.Combobox(columns_frame, textvariable=self.target_column_var, width=20, state="disabled")
        self.target_column_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(columns_frame, text="引用列:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        # 创建一个框架来容纳引用列的多选列表框和滚动条
        ref_columns_list_frame = ttk.Frame(columns_frame)
        ref_columns_list_frame.grid(row=1, column=1, sticky=tk.W+tk.E+tk.N+tk.S, padx=5, pady=5)
        
        self.ref_columns_listbox = tk.Listbox(ref_columns_list_frame, selectmode=tk.MULTIPLE, height=6, width=30)
        ref_scrollbar = ttk.Scrollbar(ref_columns_list_frame, orient=tk.VERTICAL, command=self.ref_columns_listbox.yview)
        self.ref_columns_listbox.configure(yscrollcommand=ref_scrollbar.set)
        
        self.ref_columns_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ref_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 提示语模板管理部分
        template_frame = ttk.LabelFrame(main_frame, text="提示语模板管理", padding=10)
        template_frame.pack(fill=tk.X, pady=5)
        
        # 模板选择下拉框
        template_select_frame = ttk.Frame(template_frame)
        template_select_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(template_select_frame, text="选择模板:").pack(side=tk.LEFT, padx=5)
        
        # 合并预设模板和用户模板
        all_templates = list(self.preset_templates.keys()) + list(self.templates.keys())
        
        self.template_var = tk.StringVar()
        self.template_combo = ttk.Combobox(template_select_frame, textvariable=self.template_var, width=30, state="readonly")
        self.template_combo['values'] = all_templates
        self.template_combo.pack(side=tk.LEFT, padx=5)
        self.template_combo.bind("<<ComboboxSelected>>", self.on_template_selected)
        
        # 模板操作按钮
        template_button_frame = ttk.Frame(template_frame)
        template_button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(template_button_frame, text="加载模板", 
                  command=self.load_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="保存当前模板", 
                  command=self.save_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="删除模板", 
                  command=self.delete_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="导入模板", 
                  command=self.import_templates).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="导出模板", 
                  command=self.export_templates).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="预览模板", 
                  command=self.preview_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="使用帮助", 
                  command=self.show_template_help).pack(side=tk.LEFT, padx=5)
        
        # 提示语部分
        prompt_frame = ttk.LabelFrame(main_frame, text="AI提示语", padding=10)
        prompt_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(prompt_frame, text="基于引用列提示AI生成内容:").pack(anchor=tk.W, padx=5, pady=5)
        
        # 添加提示语编辑区域
        self.prompt_text = tk.Text(prompt_frame, height=5, width=80)
        self.prompt_text.pack(fill=tk.X, padx=5, pady=5)
        self.prompt_text.insert(tk.END, "请根据以下内容生成一段总结:\n{引用内容}")
        
        # 处理按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="预览", command=self.preview_generation).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="生成并更新", command=self.generate_and_update).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出文件", command=self.export_file).pack(side=tk.LEFT, padx=5)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="预览", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        preview_scroll = ttk.Scrollbar(preview_frame)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_text = tk.Text(preview_frame, height=10, width=80, yscrollcommand=preview_scroll.set)
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        preview_scroll.config(command=self.preview_text.yview)
        
        # 将所有子控件打包完毕后，更新Canvas的滚动区域
        self.scrollable_frame.update_idletasks()  # 确保所有控件都已布局
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        
        # 窗口尺寸变化时更新Canvas窗口宽度
        self.root.bind("<Configure>", self.on_window_resize)
    
    def update_canvas_scroll_region(self):
        """更新Canvas的滚动区域"""
        self.scrollable_frame.update_idletasks()  # 确保所有控件都已布局
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
    
    def select_file(self):
        """选择Excel或CSV文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel或CSV文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            
            # 根据文件类型读取数据
            if file_path.endswith('.csv'):
                try:
                    # 先尝试使用二进制模式读取文件头部，检测BOM标记
                    with open(file_path, 'rb') as f:
                        raw_data = f.read(4)
                        
                    # 检查是否有UTF-8 BOM标记 (EF BB BF)
                    if raw_data.startswith(b'\xef\xbb\xbf'):
                        encoding = 'utf-8-sig'
                    else:
                        # 尝试通过内容检测编码
                        try:
                            import chardet
                            with open(file_path, 'rb') as f:
                                raw_data = f.read(10000)  # 读取前10KB数据用于检测
                            result = chardet.detect(raw_data)
                            encoding = result['encoding']
                            confidence = result['confidence']
                            self.status_var.set(f"检测到编码: {encoding} (置信度: {confidence:.2f})")
                        except ImportError:
                            # 如果没有安装chardet，则尝试常见编码
                            encodings_to_try = ['utf-8', 'gbk', 'cp936', 'latin1']
                            encoding = None
                            for enc in encodings_to_try:
                                try:
                                    # 尝试读取文件前几行
                                    with open(file_path, 'r', encoding=enc) as f:
                                        for _ in range(5):
                                            f.readline()
                                    encoding = enc
                                    break
                                except UnicodeDecodeError:
                                    continue
                            
                            # 如果没有找到合适的编码，默认使用utf-8
                            if not encoding:
                                encoding = 'utf-8'
                    
                    # 使用检测到的编码打开CSV文件
                    try:
                        self.df = pd.read_csv(file_path, encoding=encoding)
                        self.status_var.set(f"已使用 {encoding} 编码加载文件")
                    except Exception as e:
                        # 如果还是失败，尝试其他编码
                        success = False
                        encodings_to_try = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'cp936', 'latin1']
                        
                        # 确保不重复尝试已失败的编码
                        if encoding in encodings_to_try:
                            encodings_to_try.remove(encoding)
                        
                        for enc in encodings_to_try:
                            try:
                                self.df = pd.read_csv(file_path, encoding=enc)
                                self.status_var.set(f"已使用 {enc} 编码加载文件")
                                success = True
                                break
                            except UnicodeDecodeError:
                                continue
                            except Exception as specific_error:
                                # 记录特定编码的错误，但继续尝试其他编码
                                print(f"使用 {enc} 编码读取时出错: {str(specific_error)}")
                        
                        if not success:
                            messagebox.showerror("错误", f"无法自动检测文件编码，请使用\"手动指定编码打开\"功能。原始错误: {str(e)}")
                            self.status_var.set("读取文件失败")
                            return
                except Exception as e:
                    messagebox.showerror("错误", f"读取CSV文件时出错: {str(e)}")
                    self.status_var.set("读取文件失败")
                    return
            else:
                self.df = pd.read_excel(file_path)
            
            # 更新列选择下拉框
            self.update_column_selections()
            
            messagebox.showinfo("成功", f"成功加载文件，共 {len(self.df)} 行数据")
            self.status_var.set(f"已加载文件: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件时出错: {str(e)}")
            self.status_var.set("读取文件失败")
    
    def update_column_selections(self):
        """更新列选择下拉框和列表框"""
        if self.df is not None:
            columns = self.df.columns.tolist()
            
            # 更新目标列下拉框，添加"新建列"选项
            self.target_column_combo['values'] = ["新建列"] + columns
            self.target_column_combo['state'] = 'readonly'
            if columns:
                self.target_column_var.set(columns[0])
            else:
                self.target_column_var.set("新建列")
                
            # 绑定目标列变化事件
            self.target_column_combo.bind("<<ComboboxSelected>>", self.on_target_column_changed)
            
            # 更新引用列列表框
            self.ref_columns_listbox.delete(0, tk.END)
            for col in columns:
                self.ref_columns_listbox.insert(tk.END, col)
    
    def save_api_key(self):
        """保存API密钥和配置"""
        self.api_key = self.api_key_var.get().strip()
        self.ollama_url = self.ollama_url_var.get().strip()
        self.api_type = self.api_type_var.get()
        
        self.save_config()
        messagebox.showinfo("成功", "API配置已保存")
        
        # 如果是Ollama模式，尝试获取模型列表
        if self.api_type == "ollama":
            try:
                self.list_ollama_models(auto_update=True)
            except:
                pass
    
    def toggle_api_key_visibility(self):
        """切换API密钥的可见性"""
        if self.api_key_entry['show'] == '*':
            self.api_key_entry['show'] = ''
        else:
            self.api_key_entry['show'] = '*'
    
    def get_selected_ref_columns(self) -> List[str]:
        """获取选中的引用列"""
        selected_indices = self.ref_columns_listbox.curselection()
        if not selected_indices:
            return []
        
        selected_columns = [self.ref_columns_listbox.get(i) for i in selected_indices]
        return selected_columns
    
    def preview_generation(self):
        """预览生成的内容"""
        if self.df is None:
            messagebox.showwarning("警告", "请先选择文件")
            return
        
        target_column = self.target_column_var.get()
        ref_columns = self.get_selected_ref_columns()
        
        if not ref_columns:
            messagebox.showwarning("警告", "请选择至少一个引用列")
            return
        
        # 检查API配置
        if self.api_type_var.get() == "openai" and not self.api_key:
            messagebox.showwarning("警告", "请输入OpenAI API密钥")
            return
        elif self.api_type_var.get() == "ollama" and not self.ollama_url_var.get().strip():
            messagebox.showwarning("警告", "请输入Ollama URL")
            return
        
        # 获取提示语模板
        prompt_template = self.prompt_text.get("1.0", tk.END).strip()
        
        # 预览第一行数据
        try:
            row = self.df.iloc[0]
            ref_content = "\n".join([f"{col}: {row[col]}" for col in ref_columns])
            prompt = prompt_template.replace("{引用内容}", ref_content)
            
            # 调用API生成内容
            generated_content = self.generate_content_with_ai(ref_content)
            
            # 显示预览
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, f"引用内容:\n{ref_content}\n\n")
            self.preview_text.insert(tk.END, f"生成内容:\n{generated_content}")
            
            self.status_var.set("预览生成完成")
        except Exception as e:
            messagebox.showerror("错误", f"生成预览时出错: {str(e)}")
            self.status_var.set("预览生成失败")
    
    def generate_content_with_ai(self, reference_text):
        """使用AI生成内容"""
        if self.api_type == "openai":
            client = OpenAI(api_key=self.api_key)
            prompt = self.prompt_text.get("1.0", tk.END).strip()
            
            # 替换模板中的变量
            prompt = self.replace_template_variables(prompt, {"引用内容": reference_text})
            
            # 检查是否为医疗数据提取，特别是血压数据
            system_message = "你是一个专业的内容生成助手，仅输出最终结果，不要包含思考过程。"
            if '血压' in prompt.lower() or '收缩压' in prompt.lower() or '舒张压' in prompt.lower():
                system_message = "你是一个医疗数据提取专家。只提取最终的数据结果，不要包含任何解释、思考过程或额外文字。对于血压数据，只返回格式为'血压X/YmmHg'的结果。"
            
            try:
                response = client.chat.completions.create(
                    model=self.model_var.get(),
                    messages=[
                        {"role": "system", "content": system_message},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.7,
                )
                return self.clean_ai_output(response.choices[0].message.content)
            except Exception as e:
                return f"生成错误: {str(e)}"
        elif self.api_type == "ollama":
            url = f"{self.ollama_url.rstrip('/')}/api/chat"
            model = self.model_var.get()
            prompt = self.prompt_text.get("1.0", tk.END).strip()
            
            # 替换模板中的变量
            prompt = self.replace_template_variables(prompt, {"引用内容": reference_text})
            
            # 检查是否为deepseek模型
            is_deepseek = "deepseek" in model.lower()
            
            # 检查是否为医疗数据提取，特别是血压数据
            system_message = "你是一个专业的内容生成助手，仅输出最终结果，不要包含思考过程。"
            if '血压' in prompt.lower() or '收缩压' in prompt.lower() or '舒张压' in prompt.lower():
                system_message = "你是一个医疗数据提取专家。只提取最终的数据结果，不要包含任何解释、思考过程或额外文字。对于血压数据，只返回格式为'血压X/YmmHg'的结果。"
            
            # 尝试使用chat API
            try:
                if is_deepseek:
                    # Deepseek模型特殊处理
                    data = {
                        "model": model,
                        "messages": [
                            {"role": "system", "content": system_message},
                            {"role": "user", "content": prompt}
                        ],
                        "stream": False
                    }
                else:
                    # 常规Ollama模型
                    data = {
                        "model": model,
                        "messages": [
                            {"role": "system", "content": system_message},
                            {"role": "user", "content": prompt}
                        ],
                        "stream": False
                    }
                
                response = requests.post(url, json=data, timeout=120)
                if response.status_code == 200:
                    result = response.json()
                    if 'message' in result:
                        return self.clean_ai_output(result['message']['content'])
                    else:
                        return self.clean_ai_output(result['response'])
                else:
                    # 如果chat API失败，尝试使用generate API
                    url = f"{self.ollama_url.rstrip('/')}/api/generate"
                    data = {
                        "model": model,
                        "prompt": f"{system_message}\n\n{prompt}",
                        "stream": False
                    }
                    response = requests.post(url, json=data, timeout=120)
                    if response.status_code == 200:
                        result = response.json()
                        return self.clean_ai_output(result['response'])
                    else:
                        return f"生成错误: API返回状态码 {response.status_code}"
            except Exception as e:
                return f"生成错误: {str(e)}"
        else:
            return "错误: 未知的API类型"

    def clean_ai_output(self, text):
        """清理AI输出，去除思考过程等多余内容"""
        # 移除常见的思考过程标记
        patterns = [
            r"^(好的|嗯|我来|让我|首先|思考|理解中|分析中|处理中).*?\n",
            r"^(这是|以下是|根据|分析|结果如下).*?\n",
            r"\n*(根据提供的|参考文献|引用来源|注意事项|补充说明).*$",
        ]
        
        result = text
        for pattern in patterns:
            result = re.sub(pattern, "", result, flags=re.MULTILINE)
            
        # 去除多余的空行
        result = re.sub(r'\n{3,}', '\n\n', result)
        
        # 特殊处理血压数据提取
        # 检查是否是血压数据提取任务，如果是且找到血压格式，则只返回血压数据
        blood_pressure_pattern = r'(\d{2,3}/\d{2,3})(?:\s*(?:mmHg|毫米汞柱))'
        blood_pressure_match = re.search(blood_pressure_pattern, result)
        
        if blood_pressure_match:
            prompt = self.prompt_text.get("1.0", tk.END).strip().lower()
            if '血压' in prompt or '收缩压' in prompt or '舒张压' in prompt:
                # 如果是血压提取任务，则只返回血压数据
                bp_value = blood_pressure_match.group(1)
                return f"血压{bp_value}mmHg"
        
        return result.strip()
    
    def replace_template_variables(self, template, variables):
        """替换模板中的变量"""
        result = template
        
        # 替换简单变量 {变量名}
        for var_name, var_value in variables.items():
            pattern = r'\{' + re.escape(var_name) + r'\}'
            result = re.sub(pattern, var_value, result)
            
        # 替换带条件的变量 {如果:变量名:内容}
        for var_name, var_value in variables.items():
            # 如果变量有值，则保留条件块内容
            if var_value:
                pattern = r'\{如果:' + re.escape(var_name) + r':(.*?)\}'
                result = re.sub(pattern, r'\1', result)
            else:
                # 如果变量没有值，则移除整个条件块
                pattern = r'\{如果:' + re.escape(var_name) + r':.*?\}'
                result = re.sub(pattern, '', result)
                
        # 替换循环变量 (未来可扩展)
        
        return result

    def generate_and_update(self):
        """生成内容并更新数据"""
        if self.df is None:
            messagebox.showwarning("警告", "请先选择文件")
            return
        
        target_column = self.target_column_var.get()
        ref_columns = self.get_selected_ref_columns()
        
        if not ref_columns:
            messagebox.showwarning("警告", "请选择至少一个引用列")
            return
            
        # 检查目标列
        if target_column == "新建列":
            messagebox.showwarning("警告", "请选择或创建写入列")
            self.show_new_column_dialog()
            return
            
        # 如果是新列但不在DataFrame中，则创建它
        if target_column not in self.df.columns:
            self.df[target_column] = ""
        
        # 检查API配置
        if self.api_type_var.get() == "openai" and not self.api_key:
            messagebox.showwarning("警告", "请输入OpenAI API密钥")
            return
        elif self.api_type_var.get() == "ollama" and not self.ollama_url_var.get().strip():
            messagebox.showwarning("警告", "请输入Ollama URL")
            return
        
        # 获取提示语模板
        prompt_template = self.prompt_text.get("1.0", tk.END).strip()
        
        # 创建进度窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title("处理中")
        progress_window.geometry("450x280")
        
        ttk.Label(progress_window, text="正在生成内容，请稍候...").pack(pady=10)
        
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        
        current_label = ttk.Label(progress_window, text="0/0")
        current_label.pack(pady=5)
        
        # 添加延迟控制选项（仅对OpenAI API有用）
        delay_var = tk.BooleanVar(value=False)
        delay_frame = ttk.Frame(progress_window)
        delay_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Checkbutton(delay_frame, text="添加延迟以避免API限制(OpenAI API推荐)", 
                        variable=delay_var).pack(side=tk.LEFT)
        
        # 添加线程数量控制
        thread_frame = ttk.Frame(progress_window)
        thread_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Label(thread_frame, text="处理线程数:").pack(side=tk.LEFT)
        thread_var = tk.IntVar(value=4)  # 默认4线程
        thread_spinbox = ttk.Spinbox(thread_frame, from_=1, to=8, width=5, textvariable=thread_var)
        thread_spinbox.pack(side=tk.LEFT, padx=5)
        
        # 添加批量大小控制
        batch_frame = ttk.Frame(progress_window)
        batch_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Label(batch_frame, text="UI更新频率(行):").pack(side=tk.LEFT)
        batch_var = tk.IntVar(value=10)  # 默认每10行更新一次UI
        batch_spinbox = ttk.Spinbox(batch_frame, from_=1, to=100, width=5, textvariable=batch_var)
        batch_spinbox.pack(side=tk.LEFT, padx=5)
        
        # 启动按钮
        start_button = ttk.Button(progress_window, text="开始处理", 
                                 command=lambda: self.start_processing(
                                     progress_window, progress_var, current_label, 
                                     target_column, ref_columns, prompt_template,
                                     delay_var.get(), thread_var.get(), batch_var.get()
                                 ))
        start_button.pack(pady=10)
        
    def start_processing(self, progress_window, progress_var, current_label,
                       target_column, ref_columns, prompt_template,
                       use_delay, num_threads, batch_size):
        """启动多线程处理数据"""
        total_rows = len(self.df)
        
        # 确保目标列在DataFrame中存在
        if target_column not in self.df.columns:
            self.df[target_column] = ""  # 创建新列
        
        # 禁用启动按钮，防止重复点击
        for widget in progress_window.winfo_children():
            if isinstance(widget, ttk.Button) and widget.cget('text') == "开始处理":
                widget.configure(state="disabled")
                break
        
        # 创建结果队列和进度队列
        result_queue = queue.Queue()
        progress_queue = queue.Queue()
        
        # 创建工作线程
        threads = []
        rows_per_thread = max(1, total_rows // num_threads)
        
        # 将数据分配给各个线程
        for t in range(num_threads):
            start_idx = t * rows_per_thread
            end_idx = (t + 1) * rows_per_thread if t < num_threads - 1 else total_rows
            
            thread = threading.Thread(
                target=self.process_rows,
                args=(
                    start_idx, end_idx, ref_columns, prompt_template,
                    use_delay, result_queue, progress_queue
                )
            )
            thread.daemon = True  # 设置为守护线程
            threads.append(thread)
        
        # 启动所有工作线程
        for thread in threads:
            thread.start()
        
        # 创建更新UI的函数
        def update_ui():
            processed = 0
            results = {}
            last_content = ""
            
            try:
                # 检查进度队列和结果队列
                while processed < total_rows:
                    # 处理所有可用的进度更新
                    progress_updates = 0
                    while not progress_queue.empty() and processed < total_rows:
                        # 获取一个进度更新
                        progress_queue.get()
                        progress_updates += 1
                    
                    # 更新进度计数，确保不超过总行数
                    processed += progress_updates
                    processed = min(processed, total_rows)  # 确保不超过总行数
                    
                    # 更新进度条和标签
                    progress = (processed / total_rows) * 100
                    progress_var.set(progress)
                    current_label.config(text=f"{processed}/{total_rows}")
                    
                    # 处理所有可用的结果
                    while not result_queue.empty():
                        idx, content = result_queue.get()
                        if idx == "ERROR":
                            # 处理错误情况
                            raise Exception(content)
                        
                        results[idx] = content
                        last_content = content  # 保存最近的内容用于显示
                        
                    # 更新最新处理的内容到预览区域
                    if progress_updates > 0 and (processed % batch_size == 0 or processed == total_rows):
                        self.preview_text.delete("1.0", tk.END)
                        self.preview_text.insert(tk.END, f"已处理 {processed}/{total_rows} 行\n\n")
                        if last_content:
                            self.preview_text.insert(tk.END, f"最近一行生成内容:\n{last_content}")
                    
                    # 检查线程是否全部完成
                    active_threads = sum(1 for t in threads if t.is_alive())
                    if active_threads == 0 and len(results) >= total_rows:
                        break
                    
                    # 更新UI
                    self.root.update()
                    time.sleep(0.1)  # 短暂延迟避免UI卡顿
                
                # 将结果应用到数据框
                for idx, content in results.items():
                    self.df.at[idx, target_column] = content
                
                # 所有处理完成
                progress_window.destroy()
                messagebox.showinfo("成功", f"成功处理 {len(results)} 行数据")
                self.status_var.set(f"已完成处理 {len(results)} 行数据")
                
            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("错误", f"处理数据时出错: {str(e)}")
                self.status_var.set("处理数据失败")
        
        # 启动UI更新线程
        ui_thread = threading.Thread(target=update_ui)
        ui_thread.daemon = True
        ui_thread.start()

    def process_rows(self, start_idx, end_idx, ref_columns, prompt_template,
                   use_delay, result_queue, progress_queue):
        """处理指定范围的行（在工作线程中运行）"""
        try:
            for i in range(start_idx, end_idx):
                # 获取当前行
                index = self.df.index[i]
                row = self.df.iloc[i]
                
                # 生成提示语
                ref_content = "\n".join([f"{col}: {row[col]}" for col in ref_columns])
                prompt = prompt_template.replace("{引用内容}", ref_content)
                
                # 调用API生成内容
                generated_content = self.generate_content_with_ai(ref_content)
                
                # 将结果放入队列
                result_queue.put((index, generated_content))
                progress_queue.put(1)  # 表示完成了一行，只发送一个信号
                
                # 根据API类型和用户选择决定是否添加延迟
                if use_delay and self.api_type_var.get() == "openai":
                    time.sleep(0.5)  # 仅对OpenAI API添加延迟
                
        except Exception as e:
            # 将异常信息加入结果队列
            result_queue.put(("ERROR", str(e)))
    
    def export_file(self):
        """导出处理后的文件"""
        if self.df is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
        
        export_path = filedialog.asksaveasfilename(
            title="保存文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv")]
        )
        
        if not export_path:
            return
        
        try:
            if export_path.endswith('.csv'):
                # 允许用户选择导出编码
                encoding_window = tk.Toplevel(self.root)
                encoding_window.title("选择文件编码")
                encoding_window.geometry("300x150")
                encoding_window.transient(self.root)  # 设置为主窗口的临时窗口
                encoding_window.grab_set()  # 模态窗口
                
                ttk.Label(encoding_window, text="请选择CSV文件编码:").pack(pady=10)
                
                encoding_var = tk.StringVar(value="utf-8-sig")
                encoding_combo = ttk.Combobox(encoding_window, textvariable=encoding_var, width=20)
                encoding_combo['values'] = ["utf-8-sig", "utf-8", "gbk", "gb2312", "cp936"]
                encoding_combo.pack(pady=5)
                
                def confirm_encoding():
                    nonlocal encoding_var
                    try:
                        self.df.to_csv(export_path, index=False, encoding=encoding_var.get())
                        messagebox.showinfo("成功", f"文件已成功导出到: {export_path}")
                        self.status_var.set(f"文件已导出 (编码: {encoding_var.get()})")
                        encoding_window.destroy()
                    except Exception as e:
                        messagebox.showerror("错误", f"导出文件时出错: {str(e)}")
                        self.status_var.set("导出文件失败")
                
                ttk.Button(encoding_window, text="确定", command=confirm_encoding).pack(pady=10)
                
                # 等待窗口关闭
                self.root.wait_window(encoding_window)
            else:
                self.df.to_excel(export_path, index=False)
                messagebox.showinfo("成功", f"文件已成功导出到: {export_path}")
                self.status_var.set(f"文件已导出")
        except Exception as e:
            messagebox.showerror("错误", f"导出文件时出错: {str(e)}")
            self.status_var.set("导出文件失败")

    def select_file_with_encoding(self):
        """选择Excel或CSV文件并手动指定编码"""
        file_path = filedialog.askopenfilename(
            title="选择Excel或CSV文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        # 如果选择的是Excel文件，直接读取
        if file_path.endswith(('.xlsx', '.xls')):
            try:
                self.file_path = file_path
                self.file_label.config(text=os.path.basename(file_path))
                self.df = pd.read_excel(file_path)
                self.update_column_selections()
                messagebox.showinfo("成功", f"成功加载Excel文件，共 {len(self.df)} 行数据")
                self.status_var.set(f"已加载文件: {os.path.basename(file_path)}")
                return
            except Exception as e:
                messagebox.showerror("错误", f"读取Excel文件时出错: {str(e)}")
                self.status_var.set("读取文件失败")
                return
        
        # 对于CSV文件，打开编码选择窗口
        encoding_window = tk.Toplevel(self.root)
        encoding_window.title("选择文件编码")
        encoding_window.geometry("400x250")
        encoding_window.transient(self.root)  # 设置为主窗口的临时窗口
        encoding_window.grab_set()  # 模态窗口
        
        ttk.Label(encoding_window, text="请选择CSV文件编码:").pack(pady=10)
        
        encoding_var = tk.StringVar(value="utf-8")
        encoding_combo = ttk.Combobox(encoding_window, textvariable=encoding_var, width=20)
        encoding_combo['values'] = ["utf-8", "utf-8-sig", "gbk", "gb2312", "cp936", "latin1", "ascii"]
        encoding_combo.pack(pady=5)
        
        # 显示前几行数据的预览框
        preview_frame = ttk.LabelFrame(encoding_window, text="预览", padding=5)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)
        
        preview_scroll = ttk.Scrollbar(preview_frame)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        preview_text = tk.Text(preview_frame, height=5, width=50, yscrollcommand=preview_scroll.set)
        preview_text.pack(fill=tk.BOTH, expand=True)
        preview_scroll.config(command=preview_text.yview)
        
        def preview_with_encoding():
            """使用选定的编码预览文件内容"""
            try:
                encoding = encoding_var.get()
                # 尝试读取文件的前5行
                with open(file_path, 'r', encoding=encoding) as f:
                    lines = [next(f) for _ in range(5) if f]
                
                preview_text.delete("1.0", tk.END)
                preview_text.insert(tk.END, "".join(lines))
                
            except Exception as e:
                preview_text.delete("1.0", tk.END)
                preview_text.insert(tk.END, f"使用 {encoding} 编码预览失败: {str(e)}")
        
        ttk.Button(encoding_window, text="预览", command=preview_with_encoding).pack(pady=5)
        
        def confirm_encoding():
            try:
                encoding = encoding_var.get()
                self.file_path = file_path
                self.file_label.config(text=os.path.basename(file_path))
                
                self.df = pd.read_csv(file_path, encoding=encoding)
                self.update_column_selections()
                
                messagebox.showinfo("成功", f"成功加载CSV文件，共 {len(self.df)} 行数据")
                self.status_var.set(f"已加载文件: {os.path.basename(file_path)} (编码: {encoding})")
                encoding_window.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"使用 {encoding} 编码读取文件时出错: {str(e)}")
        
        ttk.Button(encoding_window, text="确定", command=confirm_encoding).pack(pady=10)
        
        # 等待窗口关闭
        self.root.wait_window(encoding_window)

    def show_encoding_help(self):
        """显示编码帮助信息"""
        help_window = tk.Toplevel(self.root)
        help_window.title("CSV文件编码帮助")
        help_window.geometry("600x500")
        help_window.transient(self.root)
        
        # 创建一个带滚动条的文本框
        help_frame = ttk.Frame(help_window, padding=10)
        help_frame.pack(fill=tk.BOTH, expand=True)
        
        help_scroll = ttk.Scrollbar(help_frame)
        help_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        help_text = tk.Text(help_frame, wrap=tk.WORD, yscrollcommand=help_scroll.set)
        help_text.pack(fill=tk.BOTH, expand=True)
        help_scroll.config(command=help_text.yview)
        
        # 帮助内容
        help_content = """
CSV文件编码常见问题与解决方案

1. 什么是文件编码？
   文件编码决定了文本如何以二进制形式存储在计算机中。不同的编码方式支持不同的字符集。

2. 常见的编码格式：
   - UTF-8：国际通用的编码，支持几乎所有语言的字符
   - UTF-8-SIG：带有BOM(字节顺序标记)的UTF-8
   - GBK/GB2312：中文Windows系统常用编码
   - CP936：Windows中文系统的ANSI代码页
   - Latin1：西欧语言编码
   - ASCII：基本英文字符编码

3. 为什么会出现乱码？
   当用错误的编码方式读取文件时，就会出现乱码。例如，用UTF-8尝试读取GBK编码的文件。

4. 如何解决编码问题？
   - 使用"自动检测编码"功能：程序会尝试自动识别文件编码
   - 使用"手动指定编码"功能：如果自动检测失败，可以手动选择正确的编码
   - 保存CSV时选择合适的编码：通常UTF-8-SIG是最兼容的选择

5. 中文文件编码建议：
   - 对于Excel导出的CSV文件，通常是GBK或CP936编码
   - 对于通用的CSV文件，推荐使用UTF-8编码
   - 如果需要在Excel中正确打开，建议使用UTF-8-SIG编码导出

6. 如何转换文件编码？
   - 在本程序中打开文件后导出为新文件，选择所需的编码格式
   - 使用记事本等文本编辑器打开文件，然后"另存为"时选择编码格式
   - 使用专业文本编辑器如Notepad++进行编码转换

如果您在读取文件时遇到问题，请尝试使用"手动指定编码打开"功能，并尝试不同的编码格式。
        """
        
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)  # 设置为只读
        
        # 关闭按钮
        ttk.Button(help_window, text="关闭", command=help_window.destroy).pack(pady=10)

    def on_api_type_changed(self, event):
        """处理API类型变化的回调函数"""
        self.update_api_widgets()
    
    def update_api_widgets(self):
        """根据当前选择的API类型更新显示的控件"""
        api_type = self.api_type_var.get()
        self.api_type = api_type  # 更新存储的API类型
        
        if api_type == 'openai':
            # 显示OpenAI相关控件
            self.openai_key_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
            self.api_key_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
            self.save_api_btn.grid(row=1, column=2, padx=5, pady=5)
            self.toggle_api_btn.grid(row=1, column=3, padx=5, pady=5)
            
            # 隐藏Ollama相关控件
            self.ollama_url_label.grid_forget()
            self.ollama_url_entry.grid_forget()
            self.test_ollama_btn.grid_forget()
            self.list_ollama_btn.grid_forget()
            
            # 更新模型列表
            self.model_combo['values'] = ["gpt-3.5-turbo", "gpt-4", "gpt-4-turbo"]
            if self.model_var.get() not in self.model_combo['values']:
                self.model_var.set("gpt-3.5-turbo")
                
        elif api_type == 'ollama':
            # 隐藏OpenAI相关控件
            self.openai_key_label.grid_forget()
            self.api_key_entry.grid_forget()
            self.save_api_btn.grid_forget()
            self.toggle_api_btn.grid_forget()
            
            # 显示Ollama相关控件
            self.ollama_url_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
            self.ollama_url_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
            self.test_ollama_btn.grid(row=1, column=2, padx=5, pady=5)
            self.list_ollama_btn.grid(row=1, column=3, padx=5, pady=5)
            
            # 尝试获取Ollama模型列表
            try:
                self.list_ollama_models(auto_update=True)
            except:
                # 如果获取失败，设置默认值
                self.model_combo['values'] = ["deepseek-r1:14b", "llama3", "llama2", "mistral", "gemma"]
                if self.model_var.get() not in self.model_combo['values']:
                    self.model_var.set("deepseek-r1:14b")  # 默认使用用户已安装的deepseek模型
        
        # 保存配置
        self.save_config()

    def test_ollama_connection(self):
        """测试Ollama连接"""
        try:
            response = requests.get(self.ollama_url)
            if response.status_code == 200:
                messagebox.showinfo("成功", "Ollama连接测试成功")
            else:
                messagebox.showerror("错误", f"Ollama连接测试失败，状态码: {response.status_code}")
        except Exception as e:
            messagebox.showerror("错误", f"测试Ollama连接时出错: {str(e)}")
    
    def list_ollama_models(self, auto_update=False):
        """获取Ollama模型列表"""
        try:
            ollama_url = self.ollama_url_var.get().strip()
            if not ollama_url:
                ollama_url = "http://localhost:11434"
            
            # 先尝试新的API端点
            try:    
                response = requests.get(f"{ollama_url}/api/tags")
                
                if response.status_code == 200:
                    data = response.json()
                    models = [model.get("name") for model in data.get("models", [])]
                    
                    if not models:
                        # 如果没有获取到模型，尝试兼容性的方法
                        raise Exception("未找到模型，尝试其他方法")
                else:
                    raise Exception(f"API请求失败: {response.status_code}")
                    
            except Exception as e:
                # 如果新API失败，尝试直接列出模型
                response = requests.get(f"{ollama_url}/api/models")
                
                if response.status_code == 200:
                    models_data = response.json().get("models", [])
                    models = [model.get("name") for model in models_data]
                else:
                    raise Exception(f"获取模型列表失败: {response.status_code}")
            
            # 如果列表为空，添加常用模型作为默认选项
            if not models:
                models = ["deepseek-r1:14b", "llama3", "llama2", "mistral", "gemma"]
                
            # 确保deepseek-r1:14b在列表中
            if "deepseek-r1:14b" not in models:
                models.insert(0, "deepseek-r1:14b")
                
            # 更新模型下拉框
            self.model_combo['values'] = models
            
            # 如果当前选择的模型不在列表中，选择deepseek-r1:14b或第一个可用模型
            if self.model_var.get() not in models:
                if "deepseek-r1:14b" in models:
                    self.model_var.set("deepseek-r1:14b")
                else:
                    self.model_var.set(models[0])
            
            if not auto_update:
                messagebox.showinfo("成功", f"获取到Ollama模型列表，共 {len(models)} 个模型")
            
            return models
                
        except Exception as e:
            if not auto_update:
                messagebox.showerror("错误", f"获取Ollama模型列表时出错: {str(e)}")
            
            # 设置默认模型
            default_models = ["deepseek-r1:14b", "llama3", "llama2", "mistral", "gemma"]
            self.model_combo['values'] = default_models
            
            if self.model_var.get() not in default_models:
                self.model_var.set("deepseek-r1:14b")  # 使用用户已知的模型作为默认
            
            return default_models

    def on_target_column_changed(self, event):
        """当目标列选择变化时的处理函数"""
        if self.target_column_var.get() == "新建列":
            self.show_new_column_dialog()
    
    def show_new_column_dialog(self):
        """显示新建列对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("新建列")
        dialog.geometry("300x120")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="请输入新列的名称:").pack(pady=10)
        
        new_column_var = tk.StringVar(value="AI生成内容")
        entry = ttk.Entry(dialog, textvariable=new_column_var, width=30)
        entry.pack(pady=5)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def confirm_column_name():
            new_name = new_column_var.get().strip()
            if not new_name:
                messagebox.showwarning("警告", "列名不能为空", parent=dialog)
                return
                
            # 检查列名是否已存在
            if new_name in self.df.columns:
                overwrite = messagebox.askyesno("列已存在", 
                                              f"列名 '{new_name}' 已存在，是否使用此列？", 
                                              parent=dialog)
                if not overwrite:
                    return
            
            # 更新下拉框选择
            if new_name not in self.target_column_combo['values']:
                # 确保新列名在下拉列表中可见
                values = list(self.target_column_combo['values'])
                if "新建列" in values:
                    values.remove("新建列")
                values = ["新建列"] + [new_name] + [v for v in values if v != new_name]
                self.target_column_combo['values'] = values
            
            self.target_column_var.set(new_name)
            dialog.destroy()
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="确定", command=confirm_column_name).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 按回车确认
        dialog.bind("<Return>", lambda e: confirm_column_name())

    # 以下是模板管理相关的方法
    
    def on_template_selected(self, event):
        """当从下拉框选择模板时的回调"""
        template_name = self.template_var.get()
        if not template_name:
            return
            
        # 加载选中的模板
        self.load_template()
    
    def load_template(self):
        """加载选中的模板到提示语编辑框"""
        template_name = self.template_var.get()
        if not template_name:
            messagebox.showwarning("警告", "请先选择一个模板")
            return
            
        # 从预设模板或用户模板中获取内容
        template_content = ""
        if template_name in self.preset_templates:
            template_content = self.preset_templates[template_name]
        elif template_name in self.templates:
            # 兼容旧格式的模板
            if isinstance(self.templates[template_name], str):
                template_content = self.templates[template_name]
            else:
                template_content = self.templates[template_name]["content"]
        
        if template_content:
            # 更新提示语编辑框的内容
            self.prompt_text.delete("1.0", tk.END)
            self.prompt_text.insert(tk.END, template_content)
            self.status_var.set(f"已加载模板: {template_name}")
    
    def save_template(self):
        """保存当前提示语为模板"""
        # 获取当前提示语内容
        prompt_content = self.prompt_text.get("1.0", tk.END).strip()
        if not prompt_content:
            messagebox.showwarning("警告", "提示语内容不能为空")
            return
        
        # 显示对话框让用户输入模板名称
        dialog = tk.Toplevel(self.root)
        dialog.title("保存模板")
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="请输入模板名称:").pack(pady=10)
        
        name_var = tk.StringVar()
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        # 添加模板说明输入框
        ttk.Label(dialog, text="模板说明(可选):").pack(pady=3)
        desc_var = tk.StringVar()
        desc_entry = ttk.Entry(dialog, textvariable=desc_var, width=30)
        desc_entry.pack(pady=5)
        
        def do_save():
            template_name = name_var.get().strip()
            if not template_name:
                messagebox.showwarning("警告", "模板名称不能为空", parent=dialog)
                return
                
            # 检查是否覆盖预设模板
            if template_name in self.preset_templates:
                messagebox.showwarning("警告", "不能覆盖预设模板，请使用其他名称", parent=dialog)
                return
                
            # 检查是否覆盖现有模板
            if template_name in self.templates:
                overwrite = messagebox.askyesno("确认覆盖", 
                                             f"模板 '{template_name}' 已存在，是否覆盖？", 
                                             parent=dialog)
                if not overwrite:
                    return
            
            # 准备模板数据
            template_data = {
                "content": prompt_content,
                "description": desc_var.get().strip(),
                "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # 保存模板
            self.templates[template_name] = template_data
            self.save_templates()
            
            # 更新模板下拉框
            self.update_template_combo()
            self.template_var.set(template_name)
            
            dialog.destroy()
            messagebox.showinfo("成功", f"模板 '{template_name}' 保存成功")
            self.status_var.set(f"已保存模板: {template_name}")
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="保存", command=do_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 按回车保存，Esc取消
        dialog.bind("<Return>", lambda e: do_save())
        dialog.bind("<Escape>", lambda e: dialog.destroy())
    
    def delete_template(self):
        """删除选中的模板"""
        template_name = self.template_var.get()
        if not template_name:
            messagebox.showwarning("警告", "请先选择一个模板")
            return
            
        # 不能删除预设模板
        if template_name in self.preset_templates:
            messagebox.showwarning("警告", "不能删除预设模板")
            return
            
        # 确认删除
        confirm = messagebox.askyesno("确认删除", f"确定要删除模板 '{template_name}' 吗？")
        if not confirm:
            return
            
        # 从模板字典中删除
        if template_name in self.templates:
            del self.templates[template_name]
            self.save_templates()
            
            # 更新模板下拉框
            self.update_template_combo()
            self.template_var.set("")
            
            messagebox.showinfo("成功", f"模板 '{template_name}' 已删除")
            self.status_var.set(f"已删除模板: {template_name}")
    
    def import_templates(self):
        """从JSON文件导入模板"""
        import_path = filedialog.askopenfilename(
            title="导入模板文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if not import_path:
            return
            
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                imported_templates = json.load(f)
                
            if not isinstance(imported_templates, dict):
                messagebox.showerror("错误", "导入的文件格式不正确，应为JSON字典")
                return
                
            # 显示导入确认对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("确认导入")
            dialog.geometry("500x400")
            dialog.transient(self.root)
            dialog.grab_set()
            
            ttk.Label(dialog, text="选择要导入的模板:").pack(pady=10)
            
            # 创建模板列表框和滚动条
            templates_frame = ttk.Frame(dialog)
            templates_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            
            # 创建列表视图
            columns = ("名称", "描述", "创建时间")
            template_tree = ttk.Treeview(templates_frame, columns=columns, show="headings", selectmode="extended")
            
            # 配置列标题
            template_tree.heading("名称", text="模板名称")
            template_tree.heading("描述", text="描述")
            template_tree.heading("创建时间", text="创建时间")
            
            # 配置列宽度
            template_tree.column("名称", width=150)
            template_tree.column("描述", width=200)
            template_tree.column("创建时间", width=120)
            
            # 添加滚动条
            tree_scroll = ttk.Scrollbar(templates_frame, orient=tk.VERTICAL, command=template_tree.yview)
            template_tree.configure(yscrollcommand=tree_scroll.set)
            
            template_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 填充树形视图
            for name, template_data in imported_templates.items():
                if isinstance(template_data, str):
                    # 旧格式模板
                    template_tree.insert("", tk.END, values=(name, "无描述", "未知时间"))
                else:
                    # 新格式模板
                    template_tree.insert("", tk.END, values=(
                        name, 
                        template_data.get("description", ""), 
                        template_data.get("created_at", "未知时间")
                    ))
                
            def confirm_import():
                selected_items = template_tree.selection()
                if not selected_items:
                    messagebox.showwarning("警告", "请选择至少一个模板", parent=dialog)
                    return
                    
                selected_names = [template_tree.item(item, "values")[0] for item in selected_items]
                
                # 检查是否有同名模板
                conflicts = [name for name in selected_names if name in self.templates]
                if conflicts and not messagebox.askyesno("确认覆盖", 
                                                       f"以下模板已存在，是否覆盖？\n{', '.join(conflicts)}", 
                                                       parent=dialog):
                    return
                
                # 导入选中的模板
                for name in selected_names:
                    template_data = imported_templates[name]
                    
                    # 转换旧格式模板到新格式
                    if isinstance(template_data, str):
                        template_data = {
                            "content": template_data,
                            "description": "",
                            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "imported": True
                        }
                        
                    # 标记为导入
                    if isinstance(template_data, dict) and "imported" not in template_data:
                        template_data["imported"] = True
                        
                    self.templates[name] = template_data
                
                self.save_templates()
                
                # 更新模板下拉框
                self.update_template_combo()
                
                dialog.destroy()
                messagebox.showinfo("成功", f"成功导入 {len(selected_names)} 个模板")
                self.status_var.set(f"已导入 {len(selected_names)} 个模板")
            
            def preview_selected_template():
                selected_items = template_tree.selection()
                if not selected_items:
                    messagebox.showwarning("警告", "请选择一个模板预览", parent=dialog)
                    return
                
                # 只预览第一个选中的模板
                selected_name = template_tree.item(selected_items[0], "values")[0]
                template_data = imported_templates[selected_name]
                
                # 获取模板内容
                if isinstance(template_data, str):
                    template_content = template_data
                else:
                    template_content = template_data.get("content", "")
                
                # 显示预览对话框
                TemplatePreviewDialog(dialog, template_content)
            
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="预览选中模板", 
                     command=preview_selected_template).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="导入选中项", 
                     command=confirm_import).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="全选", 
                     command=lambda: template_tree.selection_set(template_tree.get_children())).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", 
                     command=dialog.destroy).pack(side=tk.LEFT, padx=5)
                     
            # 添加键盘快捷键
            dialog.bind("<Return>", lambda e: confirm_import())
            dialog.bind("<Escape>", lambda e: dialog.destroy())
            
        except Exception as e:
            messagebox.showerror("错误", f"导入模板时出错: {str(e)}")
    
    def update_template_combo(self):
        """更新模板下拉框内容"""
        all_templates = list(self.preset_templates.keys()) + list(self.templates.keys())
        self.template_combo['values'] = all_templates
        
    def export_templates(self):
        """导出用户模板到JSON文件"""
        if not self.templates:
            messagebox.showwarning("警告", "没有用户模板可导出")
            return
            
        # 创建导出选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("选择要导出的模板")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="选择要导出的模板:").pack(pady=10)
        
        # 创建模板列表框和滚动条
        templates_frame = ttk.Frame(dialog)
        templates_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建列表视图
        columns = ("名称", "描述", "创建时间")
        template_tree = ttk.Treeview(templates_frame, columns=columns, show="headings", selectmode="extended")
        
        # 配置列标题
        template_tree.heading("名称", text="模板名称")
        template_tree.heading("描述", text="描述")
        template_tree.heading("创建时间", text="创建时间")
        
        # 配置列宽度
        template_tree.column("名称", width=150)
        template_tree.column("描述", width=200)
        template_tree.column("创建时间", width=120)
        
        # 添加滚动条
        tree_scroll = ttk.Scrollbar(templates_frame, orient=tk.VERTICAL, command=template_tree.yview)
        template_tree.configure(yscrollcommand=tree_scroll.set)
        
        template_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 填充树形视图
        for name, template_data in self.templates.items():
            if isinstance(template_data, str):
                # 旧格式模板
                template_tree.insert("", tk.END, values=(name, "无描述", "未知时间"))
            else:
                # 新格式模板
                template_tree.insert("", tk.END, values=(
                    name, 
                    template_data.get("description", ""), 
                    template_data.get("created_at", "未知时间")
                ))
            
        def export_selected():
            selected_items = template_tree.selection()
            if not selected_items:
                messagebox.showwarning("警告", "请选择至少一个模板", parent=dialog)
                return
                
            selected_names = [template_tree.item(item, "values")[0] for item in selected_items]
            
            # 选择导出文件路径
            export_path = filedialog.asksaveasfilename(
                title="导出模板",
                parent=dialog,
                defaultextension=".json",
                filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
            )
            
            if not export_path:
                return
                
            # 创建导出数据
            export_data = {}
            for name in selected_names:
                export_data[name] = self.templates[name]
                
            try:
                with open(export_path, 'w', encoding='utf-8') as f:
                    json.dump(export_data, f, ensure_ascii=False, indent=2)
                    
                dialog.destroy()
                messagebox.showinfo("成功", f"成功导出 {len(selected_names)} 个模板到 {export_path}")
                self.status_var.set(f"已导出模板到: {export_path}")
                
            except Exception as e:
                messagebox.showerror("错误", f"导出模板时出错: {str(e)}", parent=dialog)
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="导出选中项", 
                 command=export_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="全选", 
                 command=lambda: template_tree.selection_set(template_tree.get_children())).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", 
                 command=dialog.destroy).pack(side=tk.LEFT, padx=5)
                 
        # 添加键盘快捷键
        dialog.bind("<Return>", lambda e: export_selected())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def preview_template(self):
        """预览当前模板，查看变量替换效果"""
        template_content = self.prompt_text.get("1.0", tk.END).strip()
        if not template_content:
            messagebox.showwarning("警告", "提示语内容不能为空")
            return
            
        # 创建并显示预览对话框
        TemplatePreviewDialog(self.root, template_content)

    def show_template_help(self):
        """显示模板变量使用帮助"""
        help_dialog = tk.Toplevel(self.root)
        help_dialog.title("模板变量使用帮助")
        help_dialog.geometry("600x400")
        help_dialog.transient(self.root)
        
        main_frame = ttk.Frame(help_dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        help_text = """
模板变量使用帮助
==============

在创建提示词模板时，你可以使用以下变量和语法:

1. 基本变量:
   使用 {变量名} 语法引用变量。例如:
   • {引用内容} - 将被替换为选中的引用列中的内容

2. 条件变量:
   使用 {如果:变量名:内容} 语法创建条件块。
   当变量存在且非空时，显示内容。例如:
   • {如果:引用内容:请分析以下内容: {引用内容}}
   
实用技巧:
• 可以在提示词中使用多个变量
• 条件变量可以嵌套使用
• 通过"预览模板"功能可以测试模板效果

变量实例:
• 基本模板: "请根据以下内容生成摘要:\n{引用内容}"
• 条件模板: "{如果:引用内容:请分析此内容: {引用内容}}"
"""
        
        # 使用Text控件展示帮助信息，允许选择和复制
        help_scroll = ttk.Scrollbar(main_frame)
        help_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        help_text_widget = tk.Text(main_frame, wrap=tk.WORD, yscrollcommand=help_scroll.set)
        help_text_widget.pack(fill=tk.BOTH, expand=True)
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.config(state=tk.DISABLED)  # 设置为只读
        
        help_scroll.config(command=help_text_widget.yview)
        
        # 底部按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="关闭", command=help_dialog.destroy).pack()
        
        # 在对话框获取焦点
        help_dialog.focus_set()

    def on_window_resize(self, event):
        """当窗口尺寸变化时调整Canvas窗口宽度"""
        # 仅处理根窗口尺寸变化
        if event.widget == self.root:
            # 更新Canvas的滚动区域
            self.scrollable_frame.update_idletasks()
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
            
    def update_canvas_scroll_region(self):
        """更新Canvas的滚动区域"""
        self.scrollable_frame.update_idletasks()  # 确保所有控件都已布局
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

def main():
    root = tk.Tk()
    app = AIColumnGenerator(root)
    root.mainloop()

if __name__ == "__main__":
    main() 
