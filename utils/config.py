#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置管理模块

负责读取和保存配置文件
"""

import os
import configparser
from pathlib import Path


class Config:
    """配置管理类"""

    def __init__(self, config_file='config.ini'):
        """初始化配置

        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self._load_config()

    def _load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            # 如果配置文件不存在，创建默认配置
            self._create_default_config()

    def _create_default_config(self):
        """创建默认配置"""
        # 邮件配置
        self.config['Email'] = {
            'sender_email': '',
            'sender_name': 'smart',
            'password': '',
            'smtp_server': 'smtp.qiye.aliyun.com',
            'smtp_port': '465',
            'imap_server': 'imap.qiye.aliyun.com',
            'imap_port': '993',
        }
        # 模板配置
        self.config['Template'] = {
            'template_path': '',
            'email_sign': 'smart',
            'company_name': 'United Field',
        }
        # 系统设置
        self.config['Settings'] = {
            'thread_count': '3',
            'preview_count': '3',
            'enable_imap_check': 'true',
            'send_interval': '1',
        }
        # 最近文件
        self.config['LastFiles'] = {
            'last_excel': '',
            'last_template': '',
        }
        self.save()

    def get(self, section, key, fallback=''):
        """获取配置值

        Args:
            section: 配置节
            key: 配置键
            fallback: 默认值

        Returns:
            配置值
        """
        try:
            return self.config.get(section, key)
        except (configparser.NoSectionError, configparser.NoOptionError):
            return fallback

    def set(self, section, key, value):
        """设置配置值

        Args:
            section: 配置节
            key: 配置键
            value: 配置值
        """
        if section not in self.config:
            self.config.add_section(section)
        self.config.set(section, key, str(value))
        self.save()

    def save(self):
        """保存配置到文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get_email_config(self):
        """获取邮件配置"""
        return {
            'sender_email': self.get('Email', 'sender_email'),
            'sender_name': self.get('Email', 'sender_name'),
            'password': self.get('Email', 'password'),
            'smtp_server': self.get('Email', 'smtp_server'),
            'smtp_port': int(self.get('Email', 'smtp_port', '465')),
            'imap_server': self.get('Email', 'imap_server'),
            'imap_port': int(self.get('Email', 'imap_port', '993')),
        }

    def get_template_config(self):
        """获取模板配置"""
        return {
            'template_path': self.get('Template', 'template_path'),
            'email_sign': self.get('Template', 'email_sign'),
            'company_name': self.get('Template', 'company_name'),
        }

    def get_settings(self):
        """获取系统设置"""
        return {
            'thread_count': int(self.get('Settings', 'thread_count', '3')),
            'preview_count': int(self.get('Settings', 'preview_count', '3')),
            'enable_imap_check': self.get('Settings', 'enable_imap_check', 'true').lower() == 'true',
            'send_interval': int(self.get('Settings', 'send_interval', '1')),
        }
