@echo off
cd /d D:\vin-tools
rem 若端口已被占用，第二个实例会绑定失败自动退出，无副作用
"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\pythonw.exe" server_proxy.py >> server.log 2>&1
