#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Word 模板处理模块

使用 python-docx 读取 Word 模板，替换占位符，生成 HTML 邮件内容
"""

import os
import re
from docx import Document
from jinja2 import Template
from utils.logger import logger


class TemplateHandler:
    """模板处理器"""

    # 字段映射：模板变量 -> Excel 字段
    FIELD_MAP = {
        'name': 'name',
        'pay_month': 'pay_month',
        'expected_days': 'expected_days',
        'actual_days': 'actual_days',
        'base_salary': 'base_salary',
        'performance_salary': 'performance_salary',
        'service_bonus': 'service_bonus',
        'commission': 'commission',
        'live_salary': 'live_salary',
        'pre_tax_salary': 'pre_tax_salary',
        'social_security': 'social_security',
        'housing_fund': 'housing_fund',
        'current_tax': 'current_tax',
        'total_deduction': 'total_deduction',
        'net_salary': 'net_salary',
        'email_sign': 'email_sign',
        'company_name': 'company_name',
    }

    def __init__(self, template_path):
        """初始化模板处理器

        Args:
            template_path: Word 模板文件路径
        """
        self.template_path = template_path
        self.document = None
        self._load_template()

    def _load_template(self):
        """加载 Word 模板"""
        try:
            logger.info(f"正在加载 Word 模板: {self.template_path}")
            self.document = Document(self.template_path)
            logger.info("模板加载成功")
        except Exception as e:
            logger.error(f"加载 Word 模板失败: {e}")
            raise

    def render_to_html(self, employee_data, config):
        """渲染模板为 HTML

        Args:
            employee_data: 员工数据字典
            config: 模板配置（签名、公司名等）

        Returns:
            HTML 格式的邮件内容
        """
        try:
            # 准备模板变量
            template_vars = self._prepare_vars(employee_data, config)

            # 生成 HTML 内容
            html_content = self._generate_html_from_template(template_vars)

            return html_content

        except Exception as e:
            logger.error(f"渲染模板失败: {e}")
            raise

    def _prepare_vars(self, employee_data, config):
        """准备模板变量

        Args:
            employee_data: 员工数据
            config: 配置信息

        Returns:
            模板变量字典
        """
        vars_data = employee_data.copy()

        # 处理空值，显示为0或空字符串
        for key, value in vars_data.items():
            if value == '' or value is None:
                if key in ['base_salary', 'performance_salary', 'service_bonus', 'commission',
                          'live_salary', 'pre_tax_salary', 'social_security', 'housing_fund',
                          'current_tax', 'total_deduction', 'net_salary']:
                    vars_data[key] = '0'
                else:
                    vars_data[key] = ''

        # 添加签名和公司名
        vars_data['email_sign'] = config.get('email_sign', 'smart')
        vars_data['company_name'] = config.get('company_name', 'United Field')

        return vars_data

    def _generate_html_from_template(self, template_vars):
        """从模板变量生成 HTML 内容

        Args:
            template_vars: 模板变量字典

        Returns:
            HTML 格式内容
        """
        # 创建 HTML 模板 - 温馨米色调
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{
            font-family: "Microsoft YaHei UI", "微软雅黑", "SimSun", "宋体", Arial, sans-serif;
            font-size: 15px;
            line-height: 1.9;
            color: #5D4E37;
            padding: 25px;
            max-width: 680px;
            margin: 0 auto;
            background: linear-gradient(to bottom, #FFF8F0 0%, #FDF6EC 100%);
        }}
        /* 温馨卡片容器 */
        .card {{
            background: #FFFEFA;
            border-radius: 12px;
            padding: 30px;
            box-shadow: 0 2px 12px rgba(212, 165, 116, 0.15);
            border: 1px solid #F5E6D3;
        }}
        /* 问候语 */
        .greeting {{
            margin-bottom: 12px;
            color: #8B7355;
            font-size: 16px;
        }}
        /* 温馨提示条 */
        .warm-tip {{
            background: linear-gradient(to right, #FFF4E6, #FFFAF5);
            border-left: 4px solid #D4A574;
            padding: 12px 15px;
            margin: 15px 0;
            border-radius: 0 8px 8px 0;
            color: #8B7355;
            font-size: 14px;
        }}
        /* 标题 */
        .title {{
            font-size: 18px;
            font-weight: bold;
            margin: 20px 0 12px 0;
            padding-bottom: 8px;
            color: #5D4E37;
            border-bottom: 2px solid #E8D4C4;
        }}
        /* 信息行 */
        .info-row {{
            margin: 10px 0;
            color: #5D4E37;
        }}
        /* 分区标题 */
        .section-title {{
            font-weight: bold;
            margin-top: 22px;
            margin-bottom: 10px;
            color: #8B7355;
            font-size: 15px;
        }}
        /* 工资表格 */
        .salary-table {{
            border-collapse: collapse;
            width: 100%;
            margin: 12px 0;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 1px 4px rgba(212, 165, 116, 0.1);
        }}
        .salary-table td {{
            border: 1px solid #E8D4C4;
            padding: 10px 14px;
            text-align: left;
            background: #FFFEFA;
        }}
        .salary-table .header td {{
            background: linear-gradient(to bottom, #F5E6D3, #EBDCCF);
            font-weight: bold;
            text-align: center;
            color: #5D4E37;
        }}
        .salary-table tr:nth-child(even) td:not(.header) {{
            background: #FDFBF7;
        }}
        .salary-table td:first-child {{
            width: 25%;
            color: #8B7355;
        }}
        .salary-table td:nth-child(2) {{
            width: 25%;
        }}
        /* 备注区域 */
        .remarks {{
            margin-top: 18px;
            padding: 15px;
            background: #FFFAF5;
            border-radius: 8px;
            font-size: 13px;
            color: #8B7355;
            border: 1px dashed #E8D4C4;
        }}
        .remarks p {{
            margin: 6px 0;
            line-height: 1.7;
        }}
        /* 底部 */
        .footer {{
            margin-top: 25px;
            padding-top: 15px;
            color: #A89583;
            border-top: 1px solid #E8D4C4;
            text-align: right;
        }}
        .footer p {{
            margin: 5px 0;
        }}
        /* 下划线样式 */
        .underline {{
            text-decoration: underline;
            text-decoration-style: solid;
            text-decoration-color: #D4A574;
            text-decoration-thickness: 1.5px;
            padding-bottom: 1px;
        }}
        /* 数值高亮 */
        .value {{
            font-weight: 500;
            color: #8B7355;
        }}
        /* 金额强调 */
        .amount {{
            font-weight: 600;
            color: #D4A574;
            font-family: "Arial", sans-serif;
        }}
        /* 薪草装饰 */
        .decoration {{
            text-align: center;
            color: #E8D4C4;
            font-size: 24px;
            margin: 10px 0;
        }}
    </style>
</head>
<body>
    <div class="card">
        <div class="decoration">🌸 🍃 🌸</div>

        <div class="greeting">
            亲爱的 <strong class="value">{template_vars.get('name', '')}</strong>：
        </div>

        <div class="warm-tip">
            💕 温馨提示：以下是你 <strong>{template_vars.get('pay_month', '')}</strong> 的工资明细，请仔细查阅哦~
        </div>

        <div class="title">📋 工资条</div>

        <div class="info-row">
            员工姓名：<span class="value">{template_vars.get('name', '')}</span>　　　　发放月份：<span class="value">{template_vars.get('pay_month', '')}</span>
        </div>
        <div class="info-row">
            应出勤天数：<span class="value">{template_vars.get('expected_days', '')}</span> 天　　　实际出勤天数：<span class="value">{template_vars.get('actual_days', '')}</span> 天
        </div>

        <div class="section-title">💰 一、收入明细</div>
        <table class="salary-table">
            <tr class="header">
                <td>项目</td>
                <td>金额（元）</td>
                <td>项目</td>
                <td>金额（元）</td>
            </tr>
            <tr>
                <td>基本工资</td>
                <td><span class="amount underline">{template_vars.get('base_salary', '0')}</span></td>
                <td>绩效工资</td>
                <td><span class="amount underline">{template_vars.get('performance_salary', '0')}</span></td>
            </tr>
            <tr>
                <td>奖金</td>
                <td><span class="amount underline">{template_vars.get('service_bonus', '0')}</span></td>
                <td>提成</td>
                <td><span class="amount underline">{template_vars.get('commission', '0')}</span></td>
            </tr>
            <tr>
                <td>加班工资</td>
                <td><span class="amount underline">{template_vars.get('live_salary', '0')}</span></td>
                <td>其他补贴</td>
                <td><span class="amount">0</span></td>
            </tr>
        </table>

        <div class="info-row" style="margin-top: 12px;">
            <strong>应发合计：</strong><span class="amount" style="font-size: 18px; color: #C7956A;">{template_vars.get('pre_tax_salary', '0')}</span> 元
        </div>

        <div class="section-title">📝 二、扣款明细</div>
        <table class="salary-table">
            <tr class="header">
                <td>项目</td>
                <td>金额（元）</td>
                <td>项目</td>
                <td>金额（元）</td>
            </tr>
            <tr>
                <td>社保个人部分</td>
                <td><span class="amount underline">{template_vars.get('social_security', '0')}</span></td>
                <td>公积金个人部分</td>
                <td><span class="amount underline">{template_vars.get('housing_fund', '0')}</span></td>
            </tr>
            <tr>
                <td>个人所得税</td>
                <td><span class="amount underline">{template_vars.get('current_tax', '0')}</span></td>
                <td>其他扣款</td>
                <td><span class="amount">0</span></td>
            </tr>
        </table>

        <div class="info-row" style="margin-top: 12px;">
            <strong>扣款合计：</strong><span class="amount underline">{template_vars.get('total_deduction', '0')}</span> 元
        </div>

        <div class="section-title">🎁 三、实发工资</div>
        <div class="info-row" style="background: linear-gradient(to right, #FFF4E6, #FFFAF5); padding: 12px; border-radius: 8px;">
            <strong style="color: #8B7355;">实发金额：</strong><span class="amount" style="font-size: 20px; color: #C7956A;">{template_vars.get('net_salary', '0')}</span> <strong>元</strong>
        </div>

        <div class="section-title">📌 四、备注</div>
        <div class="remarks">
            <p>💡 <strong>温馨提示：</strong></p>
            <p>1. 如对工资有疑问，请随时与 HR 联系沟通~</p>
            <p>2. 工资将通过银行转账发放，请注意查收 💰</p>
            <p>3. 工资条属于个人隐私信息，请务必妥善保管 🤫</p>
        </div>

        <div class="decoration">🍂 🌿 🍂</div>

        <div class="footer">
            <p>祝您工作愉快，生活美满！✨</p>
            <p style="margin-top: 8px; color: #8B7355;">—— {template_vars.get('email_sign', 'smart')}</p>
            <p style="font-size: 13px; color: #A89583;">人力资源部</p>
        </div>
    </div>
</body>
</html>"""
        return html
