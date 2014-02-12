@echo off

"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs testBase_КонтрольФизики.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs testBase_Реиндексация.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs testBase_КонтрольЛогическойЦелостности.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs testBase_ПересчетСлужебныхДанных.ini
rem "C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_ПересчетИтогов.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs testBase_УпаковкаТаблицИБ.ini
pause