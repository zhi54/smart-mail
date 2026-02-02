#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
主窗口模块

smartMail 工资条邮件群发工具 - 主界面
现代设计风格 - 柔和优雅
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


# 现代样式配置 - 温馨优雅风格
class Styles:
    """界面样式配置 - 温馨米色暖色调"""
    # 主色调 - 温馨暖棕米色系
    PRIMARY_COLOR = "#D4A574"      # 暖金棕
    SECONDARY_COLOR = "#E8D4C4"    # 奶茶米色
    ACCENT_COLOR = "#E6B89C"       # 柔蜜桃
    HIGHLIGHT_COLOR = "#F5E6D3"    # 浅杏色

    # 渐变色
    GRADIENT_START = "#D4A574"     # 暖金棕
    GRADIENT_END = "#E8D4C4"       # 奶茶米

    # 功能色
    SUCCESS_COLOR = "#88B04B"      # 橄榄绿 (温暖绿色)
    WARNING_COLOR = "#F4C430"      # 藏红花黄
    DANGER_COLOR = "#E07A5F"       # 柔陶红

    # 背景色
    BG_COLOR = "#F9F6F0"           # 米白背景 (护眼)
    CARD_BG = "#FFFEFA"            # 奶油白卡片
    CARD_ALT_BG = "#FDFBF7"        # 交替背景

    # 文字色 - 使用棕色系代替黑色
    TEXT_COLOR = "#5D4E37"         # 深棕咖啡 (柔和不刺眼)
    TEXT_SECONDARY = "#8B7355"     # 棕褐灰
    TEXT_LIGHT = "#A89583"         # 浅棕灰

    # 边框色
    BORDER_COLOR = "#EBE0D6"       # 浅米边框
    SHADOW_COLOR = "#E8DCCF"       # 柔和阴影

    # 字体
    FONT_FAMILY = "\"Microsoft YaHei UI\", \"微软雅黑\", \"SimHei\", sans-serif"
    FONT_SIZE = 10
    FONT_LARGE = 12
    FONT_SMALL = 9


class MainWindow(tk.Tk):
    """主窗口"""

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
        self.pay_month_editable = False

        # 变量
        self.excel_path = tk.StringVar(value=self.app_config.get('LastFiles', 'last_excel'))
        self.template_path = tk.StringVar(value=self.app_config.get('LastFiles', 'last_template'))

        # 邮件配置变量
        self.sender_email = tk.StringVar(value=self.app_config.get('Email', 'sender_email'))
        self.sender_name = tk.StringVar(value=self.app_config.get('Email', 'sender_name'))
        self.email_password = tk.StringVar(value='')
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

        # 设置样式
        self._setup_styles()

        # 创建界面
        self._create_menu()
        self._create_ui()

        # 加载上次文件
        if self.excel_path.get() and os.path.exists(self.excel_path.get()):
            self._load_excel()
        if self.template_path.get() and os.path.exists(self.template_path.get()):
            self._load_template()

    def _setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()

        # 设置主题
        try:
            style.theme_use('clam')
        except:
            pass

        # 配置主框架样式
        style.configure('Card.TFrame', background=Styles.CARD_BG, relief='flat')
        style.configure('Card.TLabel', background=Styles.CARD_BG, foreground=Styles.TEXT_COLOR)
        style.configure('Card.TLabelFrame', background=Styles.CARD_BG, borderwidth=1, relief='solid')
        style.configure('Card.TLabelFrame.Label', background=Styles.CARD_BG, foreground=Styles.TEXT_COLOR, font=('Microsoft YaHei UI', 10, 'bold'))

        # 按钮样式 - 柔和圆角风格
        style.configure('Primary.TButton',
                       font=('Microsoft YaHei UI', 9, 'bold'),
                       padding=8,
                       relief='flat',
                       background=Styles.PRIMARY_COLOR,
                       foreground='white')
        style.map('Primary.TButton',
                  background=[('active', Styles.ACCENT_COLOR),
                             ('pressed', Styles.ACCENT_COLOR)])

        # 次要按钮样式
        style.configure('Secondary.TButton',
                       font=('Microsoft YaHei UI', 9),
                       padding=6,
                       relief='flat',
                       background=Styles.SECONDARY_COLOR,
                       foreground=Styles.TEXT_COLOR)
        style.map('Secondary.TButton',
                  background=[('active', Styles.PRIMARY_COLOR)])

        # 成功按钮样式
        style.configure('Success.TButton',
                       font=('Microsoft YaHei UI', 10, 'bold'),
                       padding=10,
                       relief='flat',
                       background=Styles.SUCCESS_COLOR,
                       foreground='white')
        style.map('Success.TButton',
                  background=[('active', '#7BC4B5'),
                             ('pressed', '#6AB0A3')])

        # 危险按钮样式
        style.configure('Danger.TButton',
                       font=('Microsoft YaHei UI', 10, 'bold'),
                       padding=10,
                       relief='flat',
                       background=Styles.DANGER_COLOR,
                       foreground='white')
        style.map('Danger.TButton',
                  background=[('active', '#FFA5A0'),
                             ('pressed', '#FF9390')])

        # 进度条样式 - 柔和渐变效果
        style.configure('Progress.Horizontal.TProgressbar',
                       thickness=12,
                       troughcolor=Styles.BORDER_COLOR,
                       background=Styles.PRIMARY_COLOR,
                       borderwidth=0,
                       relief='flat')

        # Treeview 样式
        style.configure('Employee.Treeview',
                       font=('Microsoft YaHei UI', 9),
                       rowheight=28,
                       background='white',
                       foreground=Styles.TEXT_COLOR,
                       fieldbackground='white',
                       borderwidth=0)
        style.configure('Employee.Treeview.Heading',
                       font=('Microsoft YaHei UI', 9, 'bold'),
                       background=Styles.HIGHLIGHT_COLOR,
                       foreground=Styles.TEXT_COLOR,
                       borderwidth=0,
                       relief='flat')
        style.map('Employee.Treeview',
                  background=[('selected', Styles.PRIMARY_COLOR)],
                  foreground=[('selected', 'white')])
        style.map('Employee.Treeview.Heading',
                  background=[('active', Styles.ACCENT_COLOR)])

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
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="退出", command=self.quit)

        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="设置", menu=settings_menu)
        settings_menu.add_command(label="邮箱配置", command=self._show_email_settings)
        settings_menu.add_command(label="系统设置", command=self._show_system_settings)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self._show_about)

    def _create_ui(self):
        """创建主界面"""
        # 主容器 - 使用灰色背景
        main_container = tk.Frame(self, bg=Styles.BG_COLOR)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 顶部标题栏
        self._create_header(main_container)

        # 内容区域
        content_frame = tk.Frame(main_container, bg=Styles.BG_COLOR)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

        # 左侧面板 - 文件选择和员工列表
        left_panel = tk.Frame(content_frame, bg=Styles.BG_COLOR)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 文件和月份选择卡片
        self._create_file_card(left_panel)

        # 员工列表卡片
        self._create_employee_card(left_panel)

        # 右侧面板 - 预览和操作
        right_panel = tk.Frame(content_frame, bg=Styles.BG_COLOR, width=680)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(10, 0))
        right_panel.pack_propagate(False)

        # 邮件预览卡片
        self._create_preview_card(right_panel)

        # 操作按钮卡片
        self._create_action_card(right_panel)

        # 状态栏
        self._create_status_bar(main_container)

    def _create_header(self, parent):
        """创建顶部标题栏 - 渐变效果"""
        # 使用 Canvas 创建渐变效果
        header_canvas = tk.Canvas(parent, height=60, highlightthickness=0)
        header_canvas.pack(fill=tk.X)

        # 绘制渐变背景 (从左到右的粉紫渐变)
        width = 1400  # 窗口宽度
        for i in range(100):
            # 计算渐变色
            r1, g1, b1 = int(Styles.PRIMARY_COLOR[1:3], 16), int(Styles.PRIMARY_COLOR[3:5], 16), int(Styles.PRIMARY_COLOR[5:7], 16)
            r2, g2, b2 = int(Styles.SECONDARY_COLOR[1:3], 16), int(Styles.SECONDARY_COLOR[3:5], 16), int(Styles.SECONDARY_COLOR[5:7], 16)

            ratio = i / 100
            r = int(r1 + (r2 - r1) * ratio)
            g = int(g1 + (g2 - g1) * ratio)
            b = int(b1 + (b2 - b1) * ratio)
            color = f"#{r:02x}{g:02x}{b:02x}"

            x0 = (width * i) // 100
            x1 = (width * (i + 1)) // 100
            header_canvas.create_rectangle(x0, 0, x1, 60, fill=color, outline="")

        # 绘制装饰性圆圈
        header_canvas.create_oval(width-150, 10, width-50, 110, fill=Styles.ACCENT_COLOR, stipple='gray25', outline="")

        # 标题
        title = tk.Label(
            header_canvas,
            text="✨ smartMail - 工资条邮件群发工具",
            bg=Styles.PRIMARY_COLOR,
            fg="white",
            font=('Microsoft YaHei UI', 15, 'bold')
        )
        title.place(x=20, y=15)

        # 右侧提示
        tips = tk.Label(
            header_canvas,
            text="💡 提示: 首次使用请先配置邮箱 → 设置 → 邮箱配置",
            bg=Styles.PRIMARY_COLOR,
            fg="white",
            font=('Microsoft YaHei UI', 9)
        )
        tips.place(x=500, y=20)

        # 绑定窗口大小变化事件
        def on_configure(event):
            header_canvas.delete("all")
            # 重新绘制渐变
            for i in range(100):
                r1, g1, b1 = int(Styles.PRIMARY_COLOR[1:3], 16), int(Styles.PRIMARY_COLOR[3:5], 16), int(Styles.PRIMARY_COLOR[5:7], 16)
                r2, g2, b2 = int(Styles.SECONDARY_COLOR[1:3], 16), int(Styles.SECONDARY_COLOR[3:5], 16), int(Styles.SECONDARY_COLOR[5:7], 16)

                ratio = i / 100
                r = int(r1 + (r2 - r1) * ratio)
                g = int(g1 + (g2 - g1) * ratio)
                b = int(b1 + (b2 - b1) * ratio)
                color = f"#{r:02x}{g:02x}{b:02x}"

                x0 = (event.width * i) // 100
                x1 = (event.width * (i + 1)) // 100
                header_canvas.create_rectangle(x0, 0, x1, 60, fill=color, outline="")

            # 重新绘制装饰
            header_canvas.create_oval(event.width-150, 10, event.width-50, 110, fill=Styles.ACCENT_COLOR, stipple='gray25', outline="")
            title.place(x=20, y=15)
            tips.place(x=event.width-400, y=20)

        # 注意：在 Tkinter 中需要绑定父窗口的 configure 事件，这里简化处理
        # 实际使用固定宽度渐变也足够美观

    def _create_file_card(self, parent):
        """创建文件选择卡片"""
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief='flat', bd=0)
        card.pack(fill=tk.X, pady=(0, 10))

        # 卡片标题
        title_frame = tk.Frame(card, bg=Styles.CARD_BG)
        title_frame.pack(fill=tk.X, padx=15, pady=(12, 8))

        tk.Label(
            title_frame,
            text="📁 数据文件",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=('Microsoft YaHei UI', 11, 'bold')
        ).pack(side=tk.LEFT)

        # 文件选择区域
        content_frame = tk.Frame(card, bg=Styles.CARD_BG)
        content_frame.pack(fill=tk.X, padx=15, pady=(0, 12))

        # Excel 文件
        row1 = tk.Frame(content_frame, bg=Styles.CARD_BG)
        row1.pack(fill=tk.X, pady=(0, 8))

        tk.Label(row1, text="Excel文件:", bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY, width=10, anchor='w').pack(side=tk.LEFT)
        excel_entry = tk.Entry(row1, textvariable=self.excel_path, bg='white', relief='flat', bd=0, font=('Microsoft YaHei UI', 9))
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), ipady=3)
        tk.Button(row1, text="📂 浏览", command=self._select_excel,
                 bg=Styles.PRIMARY_COLOR, fg='white', font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=4, borderwidth=0).pack(side=tk.LEFT)

        # 模板文件
        row2 = tk.Frame(content_frame, bg=Styles.CARD_BG)
        row2.pack(fill=tk.X, pady=(0, 8))

        tk.Label(row2, text="模板文件:", bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY, width=10, anchor='w').pack(side=tk.LEFT)
        template_entry = tk.Entry(row2, textvariable=self.template_path, bg='white', relief='flat', bd=0, font=('Microsoft YaHei UI', 9))
        template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), ipady=3)
        tk.Button(row2, text="📂 浏览", command=self._select_template,
                 bg=Styles.PRIMARY_COLOR, fg='white', font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=4, borderwidth=0).pack(side=tk.LEFT)

        # 发放月份
        row3 = tk.Frame(content_frame, bg=Styles.CARD_BG)
        row3.pack(fill=tk.X)

        tk.Label(row3, text="发放月份:", bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY, width=10, anchor='w').pack(side=tk.LEFT)

        # 月份显示
        self.pay_month_display = tk.StringVar(value="未加载")
        self.pay_month_entry = tk.Entry(row3, textvariable=self.pay_month_display, bg='white', relief='flat', bd=0, width=15, state='readonly', font=('Microsoft YaHei UI', 9))
        self.pay_month_entry.pack(side=tk.LEFT, padx=(0, 8), ipady=3)

        # 编辑按钮
        tk.Button(row3, text="✏️ 修改", command=self._edit_pay_month,
                 bg=Styles.SECONDARY_COLOR, fg=Styles.TEXT_COLOR, font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=4, borderwidth=0).pack(side=tk.LEFT)

        # 提示标签
        tk.Label(
            row3,
            text="(从Excel读取，如无该列则自动添加)",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=('Microsoft YaHei UI', 8)
        ).pack(side=tk.LEFT, padx=(10, 0))

        # 分隔线
        ttk.Separator(card, orient='horizontal').pack(fill=tk.X, padx=15, pady=5)

        # 快速统计信息
        stats_frame = tk.Frame(card, bg=Styles.CARD_BG)
        stats_frame.pack(fill=tk.X, padx=15, pady=(0, 12))

        self.stats_label = tk.Label(
            stats_frame,
            text="📊 待加载: 0 人",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=('Microsoft YaHei UI', 9)
        )
        self.stats_label.pack(side=tk.LEFT)

    def _create_employee_card(self, parent):
        """创建员工列表卡片"""
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief='flat', bd=0)
        card.pack(fill=tk.BOTH, expand=True)

        # 卡片标题和工具栏
        title_frame = tk.Frame(card, bg=Styles.CARD_BG)
        title_frame.pack(fill=tk.X, padx=15, pady=(12, 8))

        tk.Label(
            title_frame,
            text="👥 员工列表",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=('Microsoft YaHei UI', 11, 'bold')
        ).pack(side=tk.LEFT)

        # 工具按钮
        toolbar = tk.Frame(title_frame, bg=Styles.CARD_BG)
        toolbar.pack(side=tk.RIGHT)

        tk.Button(toolbar, text="✓ 全选", command=self._toggle_select_all,
                 bg=Styles.PRIMARY_COLOR, fg='white', font=('Microsoft YaHei UI', 8),
                 relief='flat', cursor='hand2', padx=10, pady=3, borderwidth=0).pack(side=tk.LEFT, padx=(0, 8))
        self.count_label = tk.Label(toolbar, text="0 人", bg=Styles.CARD_BG, fg=Styles.PRIMARY_COLOR, font=('Microsoft YaHei UI', 10, 'bold'))
        self.count_label.pack(side=tk.LEFT)

        # 列表容器
        list_container = tk.Frame(card, bg=Styles.CARD_BG)
        list_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 12))

        columns = ('select', 'name', 'email', 'pay_month', 'status')
        self.employee_tree = ttk.Treeview(list_container, columns=columns, show='headings',
                                          style='Employee.Treeview', height=12)

        self.employee_tree.heading('select', text='✓')
        self.employee_tree.heading('name', text='姓名')
        self.employee_tree.heading('email', text='邮箱')
        self.employee_tree.heading('pay_month', text='月份')
        self.employee_tree.heading('status', text='状态')

        self.employee_tree.column('select', width=35, anchor=tk.CENTER)
        self.employee_tree.column('name', width=70, anchor=tk.CENTER)
        self.employee_tree.column('email', width=180, anchor=tk.W)
        self.employee_tree.column('pay_month', width=90, anchor=tk.CENTER)
        self.employee_tree.column('status', width=55, anchor=tk.CENTER)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)

        self.employee_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.employee_tree.bind('<<TreeviewSelect>>', self._on_employee_select)

    def _create_preview_card(self, parent):
        """创建邮件预览卡片"""
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief='flat', bd=0)
        card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 卡片标题
        title_frame = tk.Frame(card, bg=Styles.CARD_BG)
        title_frame.pack(fill=tk.X, padx=15, pady=(12, 8))

        tk.Label(
            title_frame,
            text="📧 邮件预览",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=('Microsoft YaHei UI', 11, 'bold')
        ).pack(side=tk.LEFT)

        self.preview_info = tk.Label(
            title_frame,
            text="请选择员工",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=('Microsoft YaHei UI', 9)
        )
        self.preview_info.pack(side=tk.RIGHT)

        # 预览内容区域 - 限制高度，给操作卡片留空间
        preview_container = tk.Frame(card, bg=Styles.CARD_BG)
        preview_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 8))

        # 导航按钮
        nav_frame = tk.Frame(preview_container, bg=Styles.CARD_BG)
        nav_frame.pack(fill=tk.X, pady=(0, 8))

        tk.Button(nav_frame, text="◀ 上一个", command=self._prev_preview,
                 bg=Styles.SECONDARY_COLOR, fg=Styles.TEXT_COLOR, font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=5, borderwidth=0).pack(side=tk.LEFT)
        tk.Button(nav_frame, text="下一个 ▶", command=self._next_preview,
                 bg=Styles.SECONDARY_COLOR, fg=Styles.TEXT_COLOR, font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=5, borderwidth=0).pack(side=tk.LEFT, padx=(5, 0))
        tk.Button(nav_frame, text="🔄 刷新", command=self._refresh_preview,
                 bg=Styles.HIGHLIGHT_COLOR, fg=Styles.TEXT_COLOR, font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=5, borderwidth=0).pack(side=tk.LEFT, padx=(8, 0))
        tk.Button(nav_frame, text="🌐 浏览器", command=self._open_in_browser,
                 bg=Styles.PRIMARY_COLOR, fg='white', font=('Microsoft YaHei UI', 9),
                 relief='flat', cursor='hand2', padx=12, pady=5, borderwidth=0).pack(side=tk.RIGHT)

        # HTML 预览区域 - 设置最小高度，确保操作按钮可见
        preview_frame = tk.Frame(preview_container, bg='white', relief='solid', bd=1, height=350)
        preview_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 8))
        preview_frame.pack_propagate(False)

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
                self.html_frame.load_html("<html><body style='background:#f8f9fa;padding:40px;text-align:center;color:#6c757d;font-family:sans-serif;'><div style='background:white;padding:30px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);'>👈 请从左侧选择员工查看预览</div></body></html>")
            except Exception as e:
                logger.warning(f"HtmlFrame 创建失败: {e}")
                self.html_frame = None

        if self.html_frame is None:
            self.preview_text = tk.Text(
                preview_frame,
                wrap=tk.WORD,
                font=('Consolas', 9),
                bg='#fafafa'
            )
            preview_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
            self.preview_text.configure(yscrollcommand=preview_scroll.set)

            self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_action_card(self, parent):
        """创建操作按钮卡片"""
        card = tk.Frame(parent, bg=Styles.CARD_BG, relief='flat', bd=0)
        card.pack(fill=tk.X)

        # 卡片标题
        title_frame = tk.Frame(card, bg=Styles.CARD_BG)
        title_frame.pack(fill=tk.X, padx=15, pady=(12, 8))

        tk.Label(
            title_frame,
            text="🚀 发送操作",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_COLOR,
            font=('Microsoft YaHei UI', 11, 'bold')
        ).pack(side=tk.LEFT)

        # 操作按钮
        content_frame = tk.Frame(card, bg=Styles.CARD_BG)
        content_frame.pack(fill=tk.X, padx=15, pady=(0, 12))

        # 按钮行
        btn_row = tk.Frame(content_frame, bg=Styles.CARD_BG)
        btn_row.pack(fill=tk.X)

        # 开始发送按钮 - 圆角扁平风格
        self.send_btn = tk.Button(
            btn_row,
            text="💖 开始发送",
            command=self._start_send,
            bg=Styles.SUCCESS_COLOR,
            fg='white',
            font=('Microsoft YaHei UI', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=25,
            pady=10,
            borderwidth=0,
            activebackground='#7BC4B5'
        )
        self.send_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 停止按钮
        self.stop_btn = tk.Button(
            btn_row,
            text="⏹ 停止",
            command=self._stop_send,
            state=tk.DISABLED,
            bg=Styles.DANGER_COLOR,
            fg='white',
            font=('Microsoft YaHei UI', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=20,
            pady=10,
            borderwidth=0,
            activebackground='#FFA5A0'
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 导出按钮
        export_btn = tk.Button(
            btn_row,
            text="📄 导出HTML",
            command=self._export_preview,
            bg=Styles.SECONDARY_COLOR,
            fg=Styles.TEXT_COLOR,
            font=('Microsoft YaHei UI', 9),
            relief='flat',
            cursor='hand2',
            padx=15,
            pady=8,
            borderwidth=0,
            activebackground=Styles.PRIMARY_COLOR
        )
        export_btn.pack(side=tk.LEFT)

        # 进度显示
        progress_frame = tk.Frame(content_frame, bg=Styles.CARD_BG)
        progress_frame.pack(fill=tk.X, pady=(12, 0))

        # 进度条容器 - 添加圆角边框效果
        progress_container = tk.Frame(progress_frame, bg=Styles.BORDER_COLOR, padx=2, pady=2)
        progress_container.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_container,
            variable=self.progress_var,
            maximum=100,
            style='Progress.Horizontal.TProgressbar'
        )
        self.progress_bar.pack(fill=tk.X, expand=True, ipady=3)

        # 进度文本
        tk.Label(progress_frame, textvariable=self.progress_text, bg=Styles.CARD_BG,
                font=('Microsoft YaHei UI', 9), fg=Styles.TEXT_COLOR).pack(side=tk.LEFT, padx=(10, 5))
        tk.Label(progress_frame, text="•", bg=Styles.CARD_BG, fg=Styles.TEXT_SECONDARY).pack(side=tk.LEFT, padx=2)
        tk.Label(progress_frame, textvariable=self.status_text, bg=Styles.CARD_BG,
                font=('Microsoft YaHei UI', 9), fg=Styles.TEXT_SECONDARY).pack(side=tk.LEFT, padx=(5, 0))

    def _create_status_bar(self, parent):
        """创建状态栏 - 现代简洁风格"""
        status_bar = tk.Frame(parent, bg=Styles.CARD_BG, height=32)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        status_bar.pack_propagate(False)

        # 顶部装饰线
        tk.Frame(status_bar, bg=Styles.PRIMARY_COLOR, height=2).pack(fill=tk.X)

        content = tk.Frame(status_bar, bg=Styles.CARD_BG)
        content.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            content,
            text="💕  发送前请务必预览邮件内容",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=('Microsoft YaHei UI', 8)
        ).pack(side=tk.LEFT, padx=15)

        tk.Label(
            content,
            text="smartMail v1.0.0  💖",
            bg=Styles.CARD_BG,
            fg=Styles.TEXT_SECONDARY,
            font=('Microsoft YaHei UI', 8)
        ).pack(side=tk.RIGHT, padx=15)

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
            # 更新所有员工的发放月份
            for emp in self.employee_data:
                emp['pay_month'] = new_month

            # 更新显示
            self.pay_month_display.set(new_month)

            # 刷新预览
            if self.current_employee:
                self.current_employee['pay_month'] = new_month
                self._update_preview(self.current_employee)

            messagebox.showinfo("成功", f"已更新发放月份为: {new_month}")

    # ==================== 文件操作 ====================

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

            # 更新发放月份显示
            if self.employee_data:
                pay_month = self.employee_data[0].get('pay_month', '未知')
                self.pay_month_display.set(pay_month)

            # 更新统计信息
            self.stats_label.config(text=f"📊 共 {len(self.employee_data)} 人")

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

    def _load_more_employees(self):
        current_count = len(self.preview_data)
        more_count = self.settings['preview_count']
        end_index = min(current_count + more_count, len(self.employee_data))
        new_data = self.employee_data[current_count:end_index]

        if new_data:
            self.preview_data.extend(new_data)
            self._update_employee_list(self.preview_data)
            self.count_label.config(text=f"{len(self.preview_data)}/{len(self.employee_data)}")

    def _update_employee_list(self, data):
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        for employee in data:
            self.employee_tree.insert('', tk.END, values=(
                '☑',
                employee.get('name', ''),
                employee.get('email', ''),
                employee.get('pay_month', ''),
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
                self.html_frame.load_html("<html><body style='background:#f0f0f0;padding:20px;text-align:center;color:#999;'>请先选择 Word 模板文件</body></html>")
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

            # 使用 HtmlFrame 显示 HTML
            if self.html_frame:
                self.html_frame.load_html(html_content)

            # 否则显示 HTML 源码
            else:
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(1.0, html_content)

            # 更新信息
            subject = f"{employee.get('pay_month')}工资明细 - {employee.get('name')}"
            self.preview_info.config(text=f"收件: {employee.get('email')} | 主题: {subject}")

        except Exception as e:
            error_msg = f"预览生成失败：\n{e}"
            if self.html_frame:
                self.html_frame.load_html(f"<html><body style='padding:20px;color:red;'>{error_msg}</body></html>")
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

    def _open_in_browser(self):
        if not self.current_html:
            messagebox.showinfo("提示", "请先生成预览")
            return

        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(self.current_html)
                temp_path = f.name

            webbrowser.open(f'file:///{temp_path.replace(os.sep, '/')}')
            logger.info(f"在浏览器中打开预览: {temp_path}")

        except Exception as e:
            messagebox.showerror("错误", f"打开浏览器失败：\n{e}")
            logger.error(f"打开浏览器失败: {e}")

    def _refresh_preview(self):
        """刷新当前预览"""
        if self.current_employee:
            self._update_preview(self.current_employee)

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

                messagebox.showinfo("成功", f"已导出到：\n{path}")
                logger.info(f"导出预览: {path}")

            except Exception as e:
                messagebox.showerror("错误", f"导出失败：\n{e}")

    # ==================== 发送操作 ====================

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

        result = messagebox.askyesno("确认发送", f"确定要发送 {len(selected_employees)} 封邮件吗？")
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
                values[4] = '✓' if result['success'] else '✗'
                self.employee_tree.item(item, values=values)
                break

    def _on_send_complete(self):
        self.send_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_text.set("完成")

        if self.batch_sender:
            results = self.batch_sender.get_results()
            success_count = sum(1 for r in results if r['success'])
            messagebox.showinfo(
                "发送完成",
                f"共发送 {len(results)} 封\n成功: {success_count}\n失败: {len(results) - success_count}"
            )

    # ==================== 配置和设置 ====================

    def _save_config(self):
        try:
            self.app_config.set('Email', 'sender_email', self.sender_email.get())
            self.app_config.set('Email', 'sender_name', self.sender_name.get())
            self.app_config.set('Email', 'smtp_server', self.smtp_server.get())
            self.app_config.set('Email', 'smtp_port', self.smtp_port.get())
            self.app_config.set('Email', 'imap_server', self.imap_server.get())
            self.app_config.set('Email', 'imap_port', self.imap_port.get())

            password = self.email_password.get()
            if password:
                import base64
                encoded = base64.b64encode(password.encode()).decode()
                self.app_config.set('Email', 'password', encoded)

            self.app_config.set('Template', 'email_sign', self.email_sign.get())
            self.app_config.set('Template', 'company_name', self.company_name.get())

            messagebox.showinfo("成功", "配置已保存")

        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败：\n{e}")

    def _show_email_settings(self):
        SettingsDialog(self, "email")

    def _show_system_settings(self):
        SettingsDialog(self, "system")

    def _show_about(self):
        messagebox.showinfo(
            "关于",
            "smartMail 工资条邮件群发工具\n\n"
            "版本: 1.0.0\n\n"
            "功能：\n"
            "• 读取 Excel 工资数据\n"
            "• 使用 Word 模板生成邮件\n"
            "• 批量发送工资条邮件\n"
            "• 支持阿里邮箱\n\n"
            "预览：点击「浏览器中查看」查看实际邮件效果"
        )
