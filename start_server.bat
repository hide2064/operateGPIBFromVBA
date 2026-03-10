@echo off
chcp 65001 > nul
echo GPIB Flask Server を起動します...
python "%~dp0python\server.py" %*
pause
