#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
smartMail - 现代优雅风格界面
主窗口模块 - 使用 2024 流行色 Peach Fuzz + Aurora 渐变
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import tempfile
import webbrowser
from utils.config import Config
from utils.logger import logger
from core.excel_reader import ExcelReader
from core.template_handler import TemplateHandler
from core.email_sender import EmailBatchSender
from gui.preview_window import PreviewWindow
from gui.settings_dialog import SettingsDialog
from datetime import datetime
try:
    from tkinterweb import HtmlFrame
    HTMLFRAME_AVAILABLE = True
except ImportError:
    HTMLFRAME_AVAILABLE = False
    HtmlFrame = None


# 现代样式配置 - 柔和优雅风格
class Styles:
    """界面样式配置 - 2024流行色 Peach Fuzz + Aurora 渐变"""
    # 主色调 - 柔和桃粉紫渐变
    PRIMARY_COLOR = "#FFB7B2"      # 柔和桃色 (Peach Fuzz 风格)
    SECONDARY_COLOR = "#E8D5F2"   # 淡紫丁香
    ACCENT_COLOR = "#FF9EDD"      # 亮粉色
    HIGHLIGHT_COLOR = "#FFD1DC"   # 粉红高亮

    # 渐变色
    GRADIENT_START = "#FFB7CE"   # 玫瑰粉
    GRADIENT_END = "#E8D5F2"     # 淡紫

    # 功能色
    SUCCESS_COLOR = "#98D8C8"    # 柔和薄荷绿
    WARNING_COLOR = "#FFE5B4"    # 温暖橙
    DANGER_COLOR = "#FFB7B2"     # 玫瑰红

    # 背景色
    BG_COLOR = "#FFF5F7"         # 极淡粉背景
    CARD_BG = "#FFFFFF"          # 纯白卡片
    CARD_ALT_BG = "#FFFBFD"      # 交替背景

    # 文字色
    TEXT_COLOR = "#4A4A6A"      # 柔和深灰紫
    TEXT_SECONDARY = "#9B8CB8"   # 淡紫灰
    TEXT_LIGHT = "#B8A9C9"       # 浅紫灰

    # 边框色
    BORDER_COLOR = "#F0E6F0"     # 淡紫边框
    SHADOW_COLOR = "#E8D5F2"     # 柔和阴影

    # 字体
    FONT_FAMILY = "\"Microsoft YaHei UI\", \"微软雅黑\", \"SimHei\", sans-serif"
    FONT_SIZE = 10
    FONT_LARGE = 12
    FONT_SMALL = 9


def create_rounded_button(parent, text, command, bg_color, fg_color="white", width=10):
    """创建圆角按钮"""
    button = tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg_color,
        fg=fg_color,
        font=(Styles.FONT_FAMILY, Styles.FONT_SIZE, "bold"),
        relief="flat",
        cursor="hand2",
        padx=15,
        pady=8,
        borderwidth=0,
        activebackground=bg_color,
        activeforeground=fg_color
    )
    return button


def create_gradient_label(parent, text, width=400, height=60):
    """创建渐变标签"""
    canvas = tk.Canvas(parent, width=width, height=height, highlightthickness=0, bg=Styles.BG_COLOR)

    # 绘制渐变背景
    for i in range(height):
        # 计算渐变色
        ratio = i / height
        r = int(0xFF + (0xFF - 0xFF) * ratio)
        g = int(0xB7 + (0xD5 - 0xB7) * ratio)
        b = int(0xB2 + (0xF2 - 0xB2) * ratio)
        color = f"#{r:02x}{g:02x}{b:02x}"
        canvas.create_line(0, i, width, i, fill=color)

    # 添加文字
    canvas.create_text(
        width//2, height//2,
        text=text,
        fill="white",
        font=(Styles.FONT_FAMILY, Styles.FONT_LARGE, "bold")
    )

    return canvas


class MainWindow(tk.Tk):
    """主窗口 - 现代优雅设计风格"""

    def __init__(self):
        super().__init__()
        self.title("✨ smartMail - 工资条邮件群发工具")

        # 设置窗口大小和位置
        self.geometry("1400x850")
        self.minsize(1200, 750)
        self.center_window()

        # 设置窗口背景
        self.configure(bg=Styles.BG_COLOR)

        # 加载配置
        self.app_config = Config()
        self.settings = self.app_config.get_settings()

        # 数据
        self.excel_reader = None
        self.template_handler = None
        self.employee_data = []
        self.preview_data = []
        self.batch_sender = None
        self.current_html = ""
        self.current_employee = None
        self.html_frame = None

        # 变量
        self.excel_path = tk.StringVar(value=self.app_config.get('LastFiles', 'last_excel'))
        self.template_path = tk.StringVar(value=self.app_config.get('LastFiles', 'last_template'))

        # 邮件配置变量
        self.sender_email = tk.StringVar(value=self.app_config.get('Email', 'sender_email'))
        self.sender_name = tk.StringVar(value=self.app_config.get('Email', 'sender_name'))
        self.email_password = tk.StringVar(value=self.app_config.get('Email', 'password', ''))
        self.email_sign = tk.StringVar(value=self.app_config.get('Template', 'email_sign'))
        self.company_name = tk.StringVar(value=self.app_config.get('Template', 'company_name'))
        self.smtp_server = tk.StringVar(value=self.app_config.get('Email', 'smtp_server'))
        self.smtp_port = tk.StringVar(value=self.app_config.get('Email', 'smtp_port', '465'))
        self.imap_server = tk.StringVar(value=self.app_config.get('Email', 'imap_server'))
        self.imap_port = tk.StringVar(value=self.app_config.get('Email', 'imap_port', '993'))

        # 解密保存的密码
        saved_password = self.app_config.get('Email', 'password', '')
        if saved_password:
            try:
                import base64
                self.email_password.set(base64.b64decode(saved_password).decode())
            except:
                pass

        # 进度变量
        self.progress_var = tk.DoubleVar()
        self.status_text = tk.StringVar(value="就绪 ✨")
        self.progress_text = tk.StringVar(value="0/0")

        # 当前预览索引
        self.current_preview_index = 0

        # 创建界面
        self._create_menu()
        self._create_ui()

        # 加载上次文件
        if self.excel_path.get() and os.path.exists(self.excel_path.get()):
            self._load_excel()
        if self.template_path.get() and os.path.exists(self.template_path.get()):
            self._load_template()

    def center_window(self):
        """窗口居中"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self, bg=Styles.CARD_BG, fg=Styles.TEXT_COLOR)
        self.config(menu=menubar)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=False, bg=Styles.CARD_BG, fg=Styles.TEXT_COLOR)
        menubar.add_cascade(label="📁 文件", menu=file_menu)
        file_menu.add_command(label="退出", command=self.quit)

        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=False, bg=Styles.CARD_BG, fg=Styles.TEXT_COLOR)
        menubar.add_cascade(label="⚙️ 设置", menu=settings_menu)
        settings_menu.add_command(label="📧 邮箱配置", command=self._show_email_settings)
        settings_menu.add_command(label="🔧 系统设置", command=self._show_system_settings)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=False, bg=Styles.CARD_BG, fg=Styles.TEXT_COLOR)
        menubar.add_cascade(label="❓ 帮助", menu=help_menu)
        help_menu.add_command(label="ℹ️ 关于", command=self._show_about)

    def _create_ui(self):
        """创建主界面 - 现代优雅风格"""
        # 主容器
        main_container = tk.Frame(self, bg=Styles.BG_COLOR)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 顶部标题栏 - 渐变设计
        header_frame = tk.Frame(main_container, bg=Styles.BG_COLOR)
        header_frame.pack(fill=tk.X, padx=20, pady=(15, 10))

        # 渐变标题
        title_canvas = create_gradient_label(header_frame, "✨ smartMail - 工资条邮件群发工具", 600, 50)
        title_canvas.pack(side=tk.LEFT)

        # 右侧提示
        tips_label = tk.Label(
            header_frame,
            text="💖 首次使用？请先配置邮箱 → 设置 → 邮箱配置",
            bg=Styles.BG_COLOR,
            fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, 9)
        )
        tips_label.pack(side=tk.RIGHT, padx=10)

        # 内容区域
        content_frame = tk.Frame(main_container, bg=Styles.BG_COLOR)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))

        # 左侧面板
        left_panel = tk.Frame(content_frame, bg=Styles.BG_COLOR)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 右侧面板
        right_panel = tk.Frame(content_frame, bg=Styles.BG_COLOR, width=680)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(15, 0))
        right_panel.pack_propagate(False)

        # 左侧组件
        self._create_file_section(left_panel)
        self._create_employee_section(left_panel)

        # 右侧组件
        self._create_preview_section(right_panel)
        self._create_action_section(right_panel)

        # 底部状态栏
        self._create_status_bar(main_container)

    def _create_file_section(self, parent):
        """创建文件选择区域"""
        # 卡片容器 - 带阴影效果
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief="flat", bd=0)
        card.pack(fill=tk.X, pady=(0, 15))

        # 卡片内边距
        card_inner = tk.Frame(card, bg=Styles.CARD_BG)
        card_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 标题
        title = tk.Label(
            card_inner,
            text="📁 数据文件",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_LARGE, "bold")
        )
        title.pack(anchor="w", pady=(0, 12))

        # Excel 文件
        excel_frame = tk.Frame(card_inner, bg=Styles.CARD_BG)
        excel_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            excel_frame, text="📊 Excel 文件",
            bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(anchor="w")

        excel_input = tk.Frame(excel_frame, bg=Styles.CARD_BG)
        excel_input.pack(fill=tk.X, pady=(5, 0))

        tk.Entry(
            excel_input,
            textvariable=self.excel_path,
            bg="white",
            relief="solid",
            bd=1,
            highlightbackground=Styles.ACCENT_COLOR,
            highlightthickness=1,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        tk.Button(
            excel_input, text="浏览",
            command=self._select_excel,
            bg=Styles.PRIMARY_COLOR,
            fg="white",
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE, "bold"),
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=5
        ).pack(side=tk.LEFT)

        # 模板文件
        template_frame = tk.Frame(card_inner, bg=Styles.CARD_BG)
        template_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            template_frame, text="📄 模板文件",
            bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(anchor="w")

        template_input = tk.Frame(template_frame, bg=Styles.CARD_BG)
        template_input.pack(fill=tk.X, pady=(5, 0))

        tk.Entry(
            template_input,
            textvariable=self.template_path,
            bg="white",
            relief="solid",
            bd=1,
            highlightbackground=Styles.ACCENT_COLOR,
            highlightthickness=1,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        tk.Button(
            template_input, text="浏览",
            command=self._select_template,
            bg=Styles.PRIMARY_COLOR,
            fg="white",
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE, "bold"),
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=5
        ).pack(side=tk.LEFT)

        # 发放月份
        month_frame = tk.Frame(card_inner, bg=Styles.CARD_BG)
        month_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            month_frame, text="📅 发放月份",
            bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(anchor="w")

        month_input = tk.Frame(month_frame, bg=Styles.CARD_BG)
        month_input.pack(fill=tk.X, pady=(5, 0))

        self.pay_month_display = tk.StringVar(value="未加载")
        tk.Entry(
            month_input,
            textvariable=self.pay_month_display,
            bg="white",
            relief="solid",
            bd=1,
            highlightbackground=Styles.ACCENT_COLOR,
            highlightthickness=1,
            width=15,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            month_input, text="✏️",
            command=self._edit_pay_month,
            bg=Styles.WARNING_COLOR,
            fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE, "bold"),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=5
        ).pack(side=tk.LEFT)

        tk.Label(
            month_frame,
            text="(从Excel读取，如无该列则自动添加)",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_LIGHT,
            font=(Styles.FONT_FAMILY, 8)
        ).pack(side=tk.LEFT, padx=(5, 0))

    def _create_employee_section(self, parent):
        """创建员工列表区域"""
        # 卡片容器
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief="flat", bd=0)
        card.pack(fill=tk.BOTH, expand=True)

        # 卡片内边距
        card_inner = tk.Frame(card, bg=Styles.CARD_BG)
        card_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 标题栏
        title_bar = tk.Frame(card_inner, bg=Styles.CARD_BG)
        title_bar.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            title_bar,
            text="👥 员工列表",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_LARGE, "bold")
        ).pack(side=tk.LEFT)

        # 工具栏
        toolbar = tk.Frame(title_bar, bg=Styles.CARD_BG)
        toolbar.pack(side=tk.RIGHT)

        create_rounded_button(
            toolbar, "全选",
            self._toggle_select_all,
            Styles.SECONDARY_COLOR
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.count_label = tk.Label(
            toolbar,
            text="0 人",
            bg=Styles.CARD_BG,
            fg=Styles.PRIMARY_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_LARGE, "bold")
        )
        self.count_label.pack(side=tk.LEFT)

        # 列表容器
        list_container = tk.Frame(card_inner, bg=Styles.CARD_BG)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        columns = ('select', 'name', 'email', 'status')
        self.employee_tree = ttk.Treeview(
            list_container,
            columns=columns,
            show='headings',
            height=12
        )

        self.employee_tree.heading('select', text='')
        self.employee_tree.heading('name', text='姓名')
        self.employee_tree.heading('email', text='邮箱')
        self.employee_tree.heading('status', text='状态')

        self.employee_tree.column('select', width=40, anchor=tk.CENTER)
        self.employee_tree.column('name', width=80, anchor=tk.CENTER)
        self.employee_tree.column('email', width=200, anchor=tk.W)
        self.employee_tree.column('status', width=60, anchor=tk.CENTER)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)

        self.employee_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.employee_tree.bind('<<TreeviewSelect>>', self._on_employee_select)

    def _create_preview_section(self, parent):
        """创建邮件预览区域"""
        # 卡片容器
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief="flat", bd=0)
        card.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # 卡片内边距
        card_inner = tk.Frame(card, bg=Styles.CARD_BG)
        card_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 标题栏
        title_bar = tk.Frame(card_inner, bg=Styles.CARD_BG)
        title_bar.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            title_bar,
            text="📧 邮件预览",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_LARGE, "bold")
        ).pack(side=tk.LEFT)

        self.preview_info = tk.Label(
            title_bar,
            text="请选择员工",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        )
        self.preview_info.pack(side=tk.RIGHT)

        # 导航按钮
        nav_frame = tk.Frame(card_inner, bg=Styles.CARD_BG)
        nav_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Button(
            nav_frame, text="◀ 上一个",
            command=self._prev_preview,
            bg=Styles.CARD_ALT_BG, fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, 9),
            relief="flat", cursor="hand2", padx=10, pady=5
        ).pack(side=tk.LEFT)

        tk.Button(
            nav_frame, text="下一个 ▶",
            command=self._next_preview,
            bg=Styles.CARD_ALT_BG, fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, 9),
            relief="flat", cursor="hand2", padx=10, pady=5
        ).pack(side=tk.LEFT, padx=(5, 0))

        tk.Button(
            nav_frame, text="🔄",
            command=self._refresh_preview,
            bg=Styles.CARD_ALT_BG, fg=Styles.TEXT_COLOR,
            relief="flat", cursor="hand2", padx=10, pady=5
        ).pack(side=tk.LEFT, padx=(10, 0))

        tk.Button(
            nav_frame, text="🌐 浏览器",
            command=self._open_in_browser,
            bg=Styles.PRIMARY_COLOR, fg="white",
            font=(Styles.FONT_FAMILY, 9, "bold"),
            relief="flat", cursor="hand2", padx=12, pady=5
        ).pack(side=tk.RIGHT)

        # HTML 预览区域
        preview_frame = tk.Frame(card_inner, bg='white', relief='solid', bd=1)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        htmlframe_available = globals().get('HTMLFRAME_AVAILABLE', False)

        if htmlframe_available:
            try:
                self.html_frame = HtmlFrame(
                    preview_frame,
                    horizontal_scrollbar=False,
                    vertical_scrollbar=True,
                    messages_enabled=False
                )
                self.html_frame.pack(fill=tk.BOTH, expand=True)
                self.html_frame.load_html(
                    "<html><body style='background:#FFF5F7;padding:40px;text-align:center;"
                    "font-family:\"Microsoft YaHei UI\",sans-serif;color:#9B8CB8;'>"
                    "<div style='background:white;padding:30px;border-radius:12px;"
                    "box-shadow:0 4px 20px rgba(255,183,178,0.1);'>"
                    "🌸 请从左侧选择员工查看预览</div></body></html>"
                )
            except Exception as e:
                logger.warning(f"HtmlFrame 创建失败: {e}")
                self.html_frame = None

        if self.html_frame is None:
            self.preview_text = tk.Text(
                preview_frame,
                wrap=tk.WORD,
                font=('Consolas', 9),
                bg='#FFF5F7'
            )
            preview_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
            self.preview_text.configure(yscrollcommand=preview_scroll.set)
            self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_action_section(self, parent):
        """创建操作按钮区域"""
        # 卡片容器
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief="flat", bd=0)
        card.pack(fill=tk.X)

        # 卡片内边距
        card_inner = tk.Frame(card, bg=Styles.CARD_BG)
        card_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 标题
        tk.Label(
            card_inner,
            text="🚀 发送操作",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_LARGE, "bold")
        ).pack(anchor="w", pady=(0, 12))

        # 按钮行
        btn_row = tk.Frame(card_inner, bg=Styles.CARD_BG)
        btn_row.pack(fill=tk.X, pady=(0, 12))

        self.send_btn = create_rounded_button(
            btn_row, "💕 开始发送",
            self._start_send,
            Styles.SUCCESS_COLOR
        )
        self.send_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_btn = create_rounded_button(
            btn_row, "⏹ 停止",
            self._stop_send,
            Styles.DANGER_COLOR
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.stop_btn.config(state=tk.DISABLED)

        tk.Button(
            btn_row, text="📄 导出HTML",
            command=self._export_preview,
            bg=Styles.SECONDARY_COLOR, fg="white",
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE, "bold"),
            relief="flat", cursor="hand2", padx=12, pady=8
        ).pack(side=tk.LEFT)

        # 进度显示
        progress_frame = tk.Frame(card_inner, bg=Styles.CARD_BG)
        progress_frame.pack(fill=tk.X, pady=(0, 8))

        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=250,
            mode='determinate'
        )
        self.progress_bar.pack(side=tk.LEFT, padx=(0, 15))

        # 进度文本
        tk.Label(
            progress_frame,
            textvariable=self.progress_text,
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(side=tk.LEFT)

        tk.Label(
            progress_frame,
            text="|",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_LIGHT
        ).pack(side=tk.LEFT, padx=(8, 8))

        tk.Label(
            progress_frame,
            textvariable=self.status_text,
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, Styles.FONT_SIZE)
        ).pack(side=tk.LEFT)

    def _create_status_bar(self, parent):
        """创建状态栏"""
        status_bar = tk.Frame(parent, bg=Styles.CARD_BG, height=35)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        status_bar.pack_propagate(False)

        # 分隔线
        separator = tk.Frame(status_bar, bg=Styles.BORDER_COLOR, height=1)
        separator.pack(fill=tk.X)

        # 内容
        content = tk.Frame(status_bar, bg=Styles.CARD_BG)
        content.pack(fill=tk.BOTH, expand=True, padx=20)

        tk.Label(
            content,
            text="💖 提示：发送前请务必预览邮件内容",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=(Styles.FONT_FAMILY, 8)
        ).pack(side=tk.LEFT, pady=8)

        tk.Label(
            content,
            text="v1.0.0 | 💕",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_LIGHT,
            font=(Styles.FONT_FAMILY, 8)
        ).pack(side=tk.RIGHT, pady=8)

    def _edit_pay_month(self):
        """编辑发放月份"""
        if not self.employee_data:
            messagebox.showinfo("提示", "请先加载Excel文件")
            return

        current_month = self.employee_data[0].get('pay_month', '')

        from tkinter import simpledialog
        new_month = simpledialog.askstring(
            "修改发放月份",
            "请输入新的发放月份 (格式: 2025年12月):",
            initialvalue=current_month,
            parent=self
        )

        if new_month:
            for emp in self.employee_data:
                emp['pay_month'] = new_month

            self.pay_month_display.set(new_month)

            if self.current_employee:
                self.current_employee['pay_month'] = new_month
                self._update_preview(self.current_employee)

            messagebox.showinfo("成功", f"✨ 已更新发放月份为: {new_month}")

    # ==================== 保留所有原有的方法 ====================
    # (这里省略了所有原有方法，代码太长，只保留修改过的部分)

    def _select_excel(self):
        path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xls *.xlsx")])
        if path:
            self.excel_path.set(path)
            self.app_config.set('LastFiles', 'last_excel', path)
            self._load_excel()

    def _select_template(self):
        path = filedialog.askopenfilename(title="选择 Word 模板", filetypes=[("Word 文档", "*.docx")])
        if path:
            self.template_path.set(path)
            self.app_config.set('LastFiles', 'last_template', path)
            self._load_template()

    def _load_excel(self):
        try:
            path = self.excel_path.get()
            if not path or not os.path.exists(path):
                return

            logger.info(f"正在加载 Excel: {path}")
            self.excel_reader = ExcelReader(path)
            self.employee_data = self.excel_reader.get_data()
            self.preview_data = self.excel_reader.get_preview_data(self.settings['preview_count'])

            if self.employee_data:
                pay_month = self.employee_data[0].get('pay_month', '未知')
                self.pay_month_display.set(pay_month)

            self._update_employee_list(self.preview_data)
            self.count_label.config(text=f"{len(self.preview_data)}/{len(self.employee_data)}")

            if self.preview_data:
                self._update_preview(self.preview_data[0])

            logger.info(f"Excel 加载成功，共 {len(self.employee_data)} 人")

        except Exception as e:
            messagebox.showerror("错误", f"加载 Excel 失败：\n{e}")
            logger.error(f"加载 Excel 失败: {e}")

    def _load_template(self):
        try:
            path = self.template_path.get()
            if not path or not os.path.exists(path):
                return

            logger.info(f"正在加载模板: {path}")
            self.template_handler = TemplateHandler(path)
            logger.info("模板加载成功")

            if self.preview_data and self.current_preview_index < len(self.preview_data):
                self._update_preview(self.preview_data[self.current_preview_index])

        except Exception as e:
            messagebox.showerror("错误", f"加载模板失败：\n{e}")
            logger.error(f"加载模板失败: {e}")

    def _update_employee_list(self, data):
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        for employee in data:
            self.employee_tree.insert('', tk.END, values=(
                '☑',
                employee.get('name', ''),
                employee.get('email', ''),
                '待发送'
            ))

    def _on_employee_select(self, event):
        selection = self.employee_tree.selection()
        if selection:
            item = selection[0]
            values = self.employee_tree.item(item, 'values')
            name = values[1]

            for idx, emp in enumerate(self.preview_data):
                if emp.get('name') == name:
                    self.current_preview_index = idx
                    self._update_preview(emp)
                    break

    def _toggle_select_all(self):
        items = self.employee_tree.get_children()
        if not items:
            return

        first_item = self.employee_tree.item(items[0])
        is_selected = first_item['values'][0] == '☑'

        new_value = '☐' if is_selected else '☑'
        for item in items:
            values = list(self.employee_tree.item(item, 'values'))
            values[0] = new_value
            self.employee_tree.item(item, values=values)

    def _update_preview(self, employee):
        if not self.template_handler:
            if self.html_frame:
                self.html_frame.load_html("<html><body style='background:#FFF5F7;padding:40px;text-align:center;color:#9B8CB8;font-family:sans-serif;'><div style='background:white;padding:30px;border-radius:12px;box-shadow:0 4px 20px rgba(255,183,178,0.1);'>📋 请先选择 Word 模板文件</div></body></html>")
            else:
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(1.0, "请先选择 Word 模板文件")
            return

        try:
            self.current_employee = employee
            template_config = {
                'email_sign': self.email_sign.get(),
                'company_name': self.company_name.get()
            }
            html_content = self.template_handler.render_to_html(employee, template_config)
            self.current_html = html_content

            if self.html_frame:
                self.html_frame.load_html(html_content)
            else:
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(1.0, html_content)

            subject = f"{employee.get('pay_month')}工资明细 - {employee.get('name')}"
            self.preview_info.config(text=f"收件: {employee.get('email')} | 主题: {subject}")

        except Exception as e:
            error_msg = f"预览生成失败：\n{e}"
            if self.html_frame:
                self.html_frame.load_html(f"<html><body style='padding:20px;color:#FFB7B2;'>{error_msg}</body></html>")
            else:
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(1.0, error_msg)
            logger.error(f"预览生成失败: {e}")

    def _prev_preview(self):
        if self.current_preview_index > 0:
            self.current_preview_index -= 1
            self._update_preview(self.preview_data[self.current_preview_index])

    def _next_preview(self):
        if self.current_preview_index < len(self.preview_data) - 1:
            self.current_preview_index += 1
            self._update_preview(self.preview_data[self.current_preview_index])

    def _refresh_preview(self):
        if self.current_employee:
            self._update_preview(self.current_employee)

    def _open_in_browser(self):
        if not self.current_html:
            messagebox.showinfo("提示", "请先生成预览")
            return

        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(self.current_html)
                temp_path = f.name

            webbrowser.open(f'file:///{temp_path.replace(os.sep, "/")}')
            logger.info(f"在浏览器中打开预览: {temp_path}")

        except Exception as e:
            messagebox.showerror("错误", f"打开浏览器失败：\n{e}")
            logger.error(f"打开浏览器失败: {e}")

    def _export_preview(self):
        if not self.current_html:
            messagebox.showinfo("提示", "没有可导出的内容")
            return

        path = filedialog.asksaveasfilename(
            title="导出 HTML",
            defaultextension=".html",
            filetypes=[("HTML 文件", "*.html")]
        )
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(self.current_html)

                messagebox.showinfo("成功", f"✨ 已导出到：\n{path}")
                logger.info(f"导出预览: {path}")

            except Exception as e:
                messagebox.showerror("错误", f"导出失败：\n{e}")

    def _start_send(self):
        if not self.sender_email.get():
            messagebox.showerror("错误", "请输入邮箱账号")
            return

        if not self.email_password.get():
            messagebox.showerror("错误", "请输入邮箱密码")
            return

        if not self.employee_data:
            messagebox.showerror("错误", "请先加载 Excel 文件")
            return

        selected_employees = self._get_selected_employees()
        if not selected_employees:
            messagebox.showwarning("提示", "请至少选择一个员工")
            return

        result = messagebox.askyesno("确认发送", f"💕 确定要发送 {len(selected_employees)} 封邮件吗？")
        if not result:
            return

        self.send_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        email_config = {
            'sender_email': self.sender_email.get(),
            'sender_name': self.sender_name.get(),
            'password': self.email_password.get(),
            'smtp_server': self.smtp_server.get(),
            'smtp_port': int(self.smtp_port.get()),
            'imap_server': self.imap_server.get(),
            'imap_port': int(self.imap_port.get()),
            'enable_imap_check': self.settings.get('enable_imap_check', True),
            'send_interval': self.settings.get('send_interval', 1),
        }

        def send_thread():
            try:
                self.batch_sender = EmailBatchSender(
                    email_config,
                    progress_callback=self._on_send_progress
                )

                self.batch_sender.send_batch(
                    employee_list=selected_employees,
                    subject_template="{pay_month}工资明细 - {name}",
                    template_handler=self.template_handler,
                    template_config={
                        'email_sign': self.email_sign.get(),
                        'company_name': self.company_name.get()
                    }
                )

                self.after(0, lambda: self._on_send_complete())

            except Exception as e:
                self.after(0, lambda: messagebox.showerror("发送失败", str(e)))
                self.after(0, lambda: self._on_send_complete())

        threading.Thread(target=send_thread, daemon=True).start()

    def _stop_send(self):
        if self.batch_sender:
            self.batch_sender.stop()
            self.status_text.set("已停止")

    def _get_selected_employees(self):
        selected = []
        items = self.employee_tree.get_children()

        for item in items:
            values = self.employee_tree.item(item, 'values')
            if values[0] == '☑':
                name = values[1]
                for emp in self.employee_data:
                    if emp.get('name') == name:
                        selected.append(emp)
                        break

        return selected

    def _on_send_progress(self, current, total, result):
        progress = (current / total) * 100
        self.progress_var.set(progress)
        self.progress_text.set(f"{current}/{total}")
        self.status_text.set("发送中...")

        items = self.employee_tree.get_children()
        for item in items:
            values = list(self.employee_tree.item(item, 'values'))
            if values[1] == result['name']:
                values[3] = '✓' if result['success'] else '✗'
                self.employee_tree.item(item, values=values)
                break

    def _on_send_complete(self):
        self.send_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_text.set("完成 ✨")

        if self.batch_sender:
            results = self.batch_sender.get_results()
            success_count = sum(1 for r in results if r['success'])
            messagebox.showinfo(
                "发送完成",
                f"✨ 共发送 {len(results)} 封\n\n成功: {success_count} 封\n失败: {len(results) - success_count} 封"
            )

    def _show_email_settings(self):
        SettingsDialog(self, "email")

    def _show_system_settings(self):
        SettingsDialog(self, "system")

    def _show_about(self):
        messagebox.showinfo(
            "关于",
            "✨ smartMail 工资条邮件群发工具\n\n"
            "版本: 1.0.0\n\n"
            "功能：\n"
            "• 读取 Excel 工资数据\n"
            "• 使用 Word 模板生成邮件\n"
            "• 批量发送工资条邮件\n"
            "• 支持阿里邮箱\n\n"
            "界面风格：现代优雅 - Peach Fuzz 🌸"
        )
