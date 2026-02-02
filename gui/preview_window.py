#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
预览窗口模块

显示完整的邮件 HTML 预览
"""

import tkinter as tk
from tkinter import ttk
import tempfile
import os
from utils.logger import logger


class PreviewWindow(tk.Toplevel):
    """预览窗口"""

    def __init__(self, parent, employee_data, template_handler, template_config):
        super().__init__(parent)
        self.title("邮件预览")
        self.geometry("700x800")

        self.employee_data = employee_data
        self.template_handler = template_handler
        self.template_config = template_config

        self._create_ui()
        self._load_preview()

    def _create_ui(self):
        """创建界面"""
        # 信息栏
        info_frame = ttk.Frame(self, padding="10")
        info_frame.pack(fill=tk.X)

        ttk.Label(info_frame, text=f"收件人: {self.employee_data.get('email')}").pack(side=tk.LEFT)
        ttk.Label(
            info_frame,
            text=f"主题: {self.employee_data.get('pay_month')}工资明细 - {self.employee_data.get('name')}"
        ).pack(side=tk.RIGHT)

        # 分隔线
        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X)

        # 预览区域（使用 Text 组件）
        preview_frame = ttk.Frame(self, padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=scrollbar.set)

        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮栏
        btn_frame = ttk.Frame(self, padding="10")
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="在浏览器中打开", command=self._open_in_browser).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="关闭", command=self.destroy).pack(side=tk.RIGHT)

    def _load_preview(self):
        """加载预览内容"""
        try:
            html_content = self.template_handler.render_to_html(
                self.employee_data,
                self.template_config
            )

            # 简单显示（去掉 HTML 标签）
            import re
            text_content = re.sub(r'<[^>]+>', '\n', html_content)
            text_content = '\n'.join(line.strip() for line in text_content.split('\n') if line.strip())

            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, text_content)

        except Exception as e:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"预览加载失败：\n{e}")
            logger.error(f"预览加载失败: {e}")

    def _open_in_browser(self):
        """在浏览器中打开预览"""
        try:
            html_content = self.template_handler.render_to_html(
                self.employee_data,
                self.template_config
            )

            # 创建临时 HTML 文件
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_path = f.name

            # 在浏览器中打开
            import webbrowser
            webbrowser.open(f'file:///{temp_path}')

        except Exception as e:
            logger.error(f"在浏览器中打开失败: {e}")
