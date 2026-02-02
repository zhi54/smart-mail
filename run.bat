@echo off
chcp 65001 >nul
cd /d D:\stfmail

REM 激活虚拟环境
call venv\Scripts\activate

REM 运行程序
python main.py

pause
