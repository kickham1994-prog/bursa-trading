@echo off
:: 设置编码为 UTF-8，防止中文路径或输出乱码
chcp 65001 >nul

:: 切换到 Google Drive 中的项目绝对路径（/d 参数确保即使跨盘符也能正确切换）
cd /d "G:\My Drive\Bursa Trading"

:: 执行自动化总控 Python 脚本
python daily_master_pipeline.py

:: 运行完毕后暂停，方便查看运行结果（如果希望全自动无感运行，可以把下面这行删掉）
pause