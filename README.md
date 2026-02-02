# smartMail - 工资条邮件群发工具

> 温馨优雅的工资条批量发送工具，让HR工作更轻松

![Version](https://img.shields.io/badge/version-1.0.0-pink)
![Python](https://img.shields.io/badge/python-3.12+-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## ✨ 功能特点

- 📧 **批量发送**：一键批量发送工资条邮件
- 📊 **Excel支持**：支持 .xls 和 .xlsx 格式
- 📄 **Word模板**：使用 Word 模板生成精美邮件
- 👁️ **实时预览**：内置 HTML 预览，所见即所得
- 🎨 **温馨设计**：柔和米色调界面，护眼舒适
- ⚙️ **灵活配置**：支持多种邮箱服务（163、阿里云等）
- 📝 **自动月份**：自动添加或读取发放月份

## 🎨 界面预览

采用温馨米色调设计，界面柔和优雅：
- 暖金棕主色调 #D4A574
- 米白背景护眼 #F9F6F0
- 精美渐变标题栏
- 圆角卡片布局

## 📦 安装使用

### 方式一：直接使用 exe（推荐）

1. 下载 `release/smartMail.exe`
2. 双击运行即可，无需安装 Python

### 方式二：源码运行

```bash
# 克隆仓库
git clone https://github.com/zhi54/smart-mail.git
cd mart-mail

# 安装依赖
pip install -r requirements.txt

# 运行程序
python main.py
```

## 📖 使用说明

### 1. 配置邮箱

首次使用请先配置邮箱：
- 点击菜单：**设置 → 邮箱配置**
- 填写发件邮箱、密码、SMTP/IMAP 服务器
- 常见邮箱配置：
  - **163邮箱**：smtp.163.com:465, imap.163.com:993
  - **阿里云邮箱**：smtp.mxhichina.com:465, imap.mxhichina.com:993

### 2. 准备数据

**Excel 文件要求**（`工资条.xls`）：
- 必须包含列：`姓名`、`邮箱`
- 可选列：`发放月份`（如无则自动添加上个月）
- 其他列：基本工资、绩效工资、奖金等

**Word 模板**（`工资条_template.docx`）：
- 使用 `{字段名}` 作为占位符
- 程序会自动替换为实际数据

### 3. 发送流程

1. 选择 Excel 文件和 Word 模板
2. 预览邮件内容（可翻页查看）
3. 勾选要发送的员工
4. 点击 **💖 开始发送**
5. 等待发送完成，查看结果

## 📂 项目结构

```
stfmail/
├── main.py                 # 程序入口
├── requirements.txt        # 依赖列表
├── core/                   # 核心功能模块
│   ├── excel_reader.py    # Excel 读取
│   ├── template_handler.py # 模板处理
│   └── email_sender.py    # 邮件发送
├── gui/                    # 图形界面
│   ├── main_window.py     # 主窗口
│   ├── settings_dialog.py # 设置对话框
│   └── preview_window.py  # 预览窗口
└── utils/                  # 工具模块
    ├── config.py          # 配置管理
    └── logger.py          # 日志记录
```

## 🛠️ 技术栈

- **GUI框架**：Tkinter + tkinterweb
- **Excel处理**：openpyxl + xlrd
- **Word处理**：python-docx
- **邮件发送**：smtplib + imaplib
- **模板引擎**：Jinja2
- **打包工具**：PyInstaller

## 📝 邮件模板示例

程序生成的邮件包含：
- 🌸 装饰图案
- 💕 温馨提示条
- 📋 工资明细表格
- 🎁 实发金额高亮
- 📌 温馨备注

## ⚠️ 注意事项

1. **邮箱密码**：部分邮箱需要使用"授权码"而非登录密码
2. **发送间隔**：建议设置间隔避免被识别为垃圾邮件
3. **数据安全**：工资条属于敏感信息，请妥善保管
4. **测试发送**：首次使用建议先发送测试邮件

## 🔄 更新日志

### v1.0.0 (2026-02-02)
- ✨ 初始版本发布
- 📧 支持批量发送工资条邮件
- 🎨 温馨米色调界面设计
- 👁️ HTML 实时预览
- 📊 自动处理发放月份

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 💝 致谢

感谢所有使用本工具的朋友们！

---

**smartMail** - 让工资条发送更温馨 ✨
