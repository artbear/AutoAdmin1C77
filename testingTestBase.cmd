@echo off
	rem cscript 1C77_Utils.vbs update_test.ini 
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs testBase.ini
echo Результат выполнения скрипта = %errorlevel%
pause