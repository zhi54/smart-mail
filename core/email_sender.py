#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
邮件发送模块

使用 SMTP 发送邮件，使用 IMAP 验证发送状态
"""

import smtplib
import imaplib
import time
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from utils.logger import logger


class EmailSender:
    """邮件发送器"""

    def __init__(self, config):
        """初始化邮件发送器

        Args:
            config: 邮件配置字典
        """
        self.config = config
        self.smtp = None
        self.imap = None

    def connect_smtp(self):
        """连接 SMTP 服务器"""
        try:
            logger.info(f"正在连接 SMTP 服务器: {self.config['smtp_server']}:{self.config['smtp_port']}")

            if self.config['smtp_port'] == 465:
                self.smtp = smtplib.SMTP_SSL(
                    self.config['smtp_server'],
                    self.config['smtp_port'],
                    timeout=30
                )
            elif self.config['smtp_port'] == 25:
                self.smtp = smtplib.SMTP(
                    self.config['smtp_server'],
                    self.config['smtp_port'],
                    timeout=30
                )
            else:
                raise ValueError(f"不支持的 SMTP 端口: {self.config['smtp_port']}")

            # 登录
            self.smtp.login(self.config['sender_email'], self.config['password'])
            logger.info("SMTP 连接成功")
            return True

        except Exception as e:
            logger.error(f"SMTP 连接失败: {e}")
            raise

    def connect_imap(self):
        """连接 IMAP 服务器"""
        try:
            if not self.config.get('enable_imap_check', False):
                logger.info("IMAP 验证已禁用")
                return True

            logger.info(f"正在连接 IMAP 服务器: {self.config['imap_server']}:{self.config['imap_port']}")
            self.imap = imaplib.IMAP4_SSL(
                self.config['imap_server'],
                self.config['imap_port'],
                timeout=30
            )
            self.imap.login(self.config['sender_email'], self.config['password'])
            logger.info("IMAP 连接成功")
            return True

        except Exception as e:
            logger.warning(f"IMAP 连接失败: {e}")
            return False

    def send_email(self, to_email, subject, html_content, sender_name=None):
        """发送单封邮件

        Args:
            to_email: 收件人邮箱
            subject: 邮件主题
            html_content: HTML 格式的邮件内容
            sender_name: 发件人名称

        Returns:
            发送结果 (True/False)
        """
        try:
            # 创建邮件
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = formataddr([
                sender_name or self.config['sender_name'],
                self.config['sender_email']
            ])
            msg['To'] = to_email

            # 添加 HTML 内容
            html_part = MIMEText(html_content, 'html', 'utf-8')
            msg.attach(html_part)

            # 发送邮件
            self.smtp.sendmail(
                self.config['sender_email'],
                [to_email],
                msg.as_string()
            )

            logger.info(f"邮件发送成功: {to_email}")
            return True

        except Exception as e:
            logger.error(f"邮件发送失败 {to_email}: {e}")
            # 重试一次
            try:
                time.sleep(1)
                self.smtp.sendmail(
                    self.config['sender_email'],
                    [to_email],
                    msg.as_string()
                )
                logger.info(f"邮件重试发送成功: {to_email}")
                return True
            except Exception as retry_error:
                logger.error(f"邮件重试发送失败 {to_email}: {retry_error}")
                return False

    def disconnect(self):
        """断开连接"""
        try:
            if self.smtp:
                self.smtp.quit()
                logger.info("SMTP 连接已断开")
        except:
            pass

        try:
            if self.imap:
                self.imap.logout()
                logger.info("IMAP 连接已断开")
        except:
            pass


class EmailBatchSender:
    """批量邮件发送器"""

    def __init__(self, config, progress_callback=None):
        """初始化批量发送器

        Args:
            config: 邮件配置
            progress_callback: 进度回调函数
        """
        self.config = config
        self.progress_callback = progress_callback
        self.sender = EmailSender(config)
        self.is_running = False
        self.is_paused = False
        self.results = []

    def send_batch(self, employee_list, subject_template, template_handler, template_config):
        """批量发送邮件

        Args:
            employee_list: 员工数据列表
            subject_template: 邮件主题模板，如 "{name}的工资明细"
            template_handler: 模板处理器
            template_config: 模板配置

        Returns:
            发送结果列表
        """
        self.is_running = True
        self.results = []

        try:
            # 连接服务器
            self.sender.connect_smtp()
            self.sender.connect_imap()

            total = len(employee_list)
            logger.info(f"开始批量发送邮件，共 {total} 封")

            for idx, employee in enumerate(employee_list):
                if not self.is_running:
                    logger.info("发送已停止")
                    break

                while self.is_paused:
                    time.sleep(0.5)
                    if not self.is_running:
                        break

                try:
                    # 生成邮件主题
                    subject = subject_template.format(
                        name=employee.get('name', ''),
                        pay_month=employee.get('pay_month', '')
                    )

                    # 生成邮件内容
                    html_content = template_handler.render_to_html(employee, template_config)

                    # 发送邮件
                    success = self.sender.send_email(
                        to_email=employee['email'],
                        subject=subject,
                        html_content=html_content,
                        sender_name=self.config.get('sender_name')
                    )

                    # 记录结果
                    result = {
                        'name': employee.get('name'),
                        'email': employee['email'],
                        'success': success,
                        'message': '成功' if success else '失败'
                    }
                    self.results.append(result)

                    # 更新进度
                    if self.progress_callback:
                        self.progress_callback(idx + 1, total, result)

                    # 发送间隔
                    time.sleep(self.config.get('send_interval', 1))

                except Exception as e:
                    logger.error(f"发送邮件失败 {employee.get('name')}: {e}")
                    result = {
                        'name': employee.get('name'),
                        'email': employee['email'],
                        'success': False,
                        'message': str(e)
                    }
                    self.results.append(result)

                    if self.progress_callback:
                        self.progress_callback(idx + 1, total, result)

            logger.info(f"批量发送完成，成功 {sum(1 for r in self.results if r['success'])} 封")

        except Exception as e:
            logger.error(f"批量发送失败: {e}")
            raise
        finally:
            self.sender.disconnect()
            self.is_running = False

        return self.results

    def stop(self):
        """停止发送"""
        self.is_running = False

    def pause(self):
        """暂停发送"""
        self.is_paused = True

    def resume(self):
        """恢复发送"""
        self.is_paused = False

    def get_results(self):
        """获取发送结果"""
        return self.results
