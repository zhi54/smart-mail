#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
设置对话框模块
"""

import tkinter as tk
from tkinter import ttk, messagebox
from utils.config import Config


class SettingsDialog(tk.Toplevel):
    """设置对话框"""

    def __init__(self, parent, dialog_type):
        super().__init__(parent)
        self.dialog_type = dialog_type
        self.config = Config()

        if dialog_type == "email":
            self.title("邮箱配置")
            self._create_email_settings()
            self.geometry("500x450")  # 增加高度以容纳密码字段
        elif dialog_type == "system":
            self.title("系统设置")
            self._create_system_settings()
            self.geometry("500x350")
        self.transient(parent)
        self.grab_set()

    def _create_email_settings(self):
        """创建邮箱设置界面"""
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        row = 0
        self.vars = {}

        # 解密密码
        saved_password = self.config.get('Email', 'password', '')
        password_value = ''
        if saved_password:
            try:
                import base64
                password_value = base64.b64decode(saved_password).decode()
            except:
                pass

        settings = [
            ('发件邮箱', 'Email', 'sender_email', False),
            ('邮箱密码', 'Email', 'password', True),
            ('发件人名称', 'Email', 'sender_name', False),
            ('SMTP 服务器', 'Email', 'smtp_server', False),
            ('SMTP 端口', 'Email', 'smtp_port', False),
            ('IMAP 服务器', 'Email', 'imap_server', False),
            ('IMAP 端口', 'Email', 'imap_port', False),
            ('邮件签名', 'Template', 'email_sign', False),
            ('公司名称', 'Template', 'company_name', False),
        ]

        for label, section, key, is_password in settings:
            ttk.Label(frame, text=label + ":").grid(row=row, column=0, sticky=tk.W, pady=5)

            # 获取值
            if key == 'password':
                value = password_value
            else:
                value = self.config.get(section, key)

            var = tk.StringVar(value=value)
            if is_password:
                entry = ttk.Entry(frame, textvariable=var, width=30, show="*")
            else:
                entry = ttk.Entry(frame, textvariable=var, width=30)
            entry.grid(row=row, column=1, pady=5, sticky=tk.W)
            self.vars[key] = var
            row += 1

        # 按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="保存", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side=tk.LEFT)

    def _create_system_settings(self):
        """创建系统设置界面"""
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        row = 0
        self.vars = {}

        # 预览数量
        ttk.Label(frame, text="默认预览数量:").grid(row=row, column=0, sticky=tk.W, pady=5)
        preview_var = tk.IntVar(value=self.config.get('Settings', 'preview_count', '3'))
        ttk.Spinbox(frame, from_=1, to=50, textvariable=preview_var, width=10).grid(row=row, column=1, sticky=tk.W, pady=5)
        self.vars['preview_count'] = preview_var
        row += 1

        # 发送线程数
        ttk.Label(frame, text="发送线程数:").grid(row=row, column=0, sticky=tk.W, pady=5)
        thread_var = tk.IntVar(value=self.config.get('Settings', 'thread_count', '3'))
        ttk.Spinbox(frame, from_=1, to=10, textvariable=thread_var, width=10).grid(row=row, column=1, sticky=tk.W, pady=5)
        self.vars['thread_count'] = thread_var
        row += 1

        # 发送间隔
        ttk.Label(frame, text="发送间隔（秒）:").grid(row=row, column=0, sticky=tk.W, pady=5)
        interval_var = tk.IntVar(value=self.config.get('Settings', 'send_interval', '1'))
        ttk.Spinbox(frame, from_=0, to=10, textvariable=interval_var, width=10).grid(row=row, column=1, sticky=tk.W, pady=5)
        self.vars['send_interval'] = interval_var
        row += 1

        # IMAP 验证
        imap_var = tk.BooleanVar(value=self.config.get('Settings', 'enable_imap_check', 'true').lower() == 'true')
        ttk.Checkbutton(frame, text="启用 IMAP 验证", variable=imap_var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=5)
        self.vars['enable_imap_check'] = imap_var
        row += 1

        # 按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="保存", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side=tk.LEFT)

    def _save(self):
        """保存设置"""
        try:
            if self.dialog_type == "email":
                # 加密保存密码
                import base64
                password = self.vars['password'].get()
                if password:
                    encoded_password = base64.b64encode(password.encode()).decode()
                    self.config.set('Email', 'password', encoded_password)

                self.config.set('Email', 'sender_email', self.vars['sender_email'].get())
                self.config.set('Email', 'sender_name', self.vars['sender_name'].get())
                self.config.set('Email', 'smtp_server', self.vars['smtp_server'].get())
                self.config.set('Email', 'smtp_port', self.vars['smtp_port'].get())
                self.config.set('Email', 'imap_server', self.vars['imap_server'].get())
                self.config.set('Email', 'imap_port', self.vars['imap_port'].get())
                self.config.set('Template', 'email_sign', self.vars['email_sign'].get())
                self.config.set('Template', 'company_name', self.vars['company_name'].get())

            elif self.dialog_type == "system":
                self.config.set('Settings', 'preview_count', self.vars['preview_count'].get())
                self.config.set('Settings', 'thread_count', self.vars['thread_count'].get())
                self.config.set('Settings', 'send_interval', self.vars['send_interval'].get())
                self.config.set('Settings', 'enable_imap_check', str(self.vars['enable_imap_check'].get()))

            messagebox.showinfo("成功", "设置已保存")
            self.destroy()

        except Exception as e:
            messagebox.showerror("错误", f"保存失败：\n{e}")
