#!/bin/bash
# STFMail 启动脚本（使用虚拟环境）

cd /d/stfmail

# 激活虚拟环境
source venv/Scripts/activate

# 运行程序
python main.py
