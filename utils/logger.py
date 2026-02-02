#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
日志记录模块
"""

import logging
import os
from datetime import datetime


class Logger:
    """日志记录类"""

    def __init__(self, log_dir='logs', log_level=logging.INFO):
        """初始化日志

        Args:
            log_dir: 日志目录
            log_level: 日志级别
        """
        self.log_dir = log_dir
        self.log_level = log_level
        self._setup_logger()

    def _setup_logger(self):
        """设置日志记录器"""
        # 创建日志目录
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)

        # 创建日志文件名（按日期）
        log_file = os.path.join(
            self.log_dir,
            f"stfmail_{datetime.now().strftime('%Y%m%d')}.log"
        )

        # 配置日志格式
        logging.basicConfig(
            level=self.log_level,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )

        self.logger = logging.getLogger('STFMail')

    def info(self, message):
        """记录信息日志"""
        self.logger.info(message)

    def error(self, message):
        """记录错误日志"""
        self.logger.error(message)

    def warning(self, message):
        """记录警告日志"""
        self.logger.warning(message)

    def debug(self, message):
        """记录调试日志"""
        self.logger.debug(message)


# 创建全局日志实例
logger = Logger()
