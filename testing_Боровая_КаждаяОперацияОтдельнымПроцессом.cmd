@echo off

"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_КонтрольФизики.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_Реиндексация.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_КонтрольЛогическойЦелостности.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_ПересчетСлужебныхДанных.ini
rem "C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_ПересчетИтогов.ini
"C:\Windows\SysWOW64\cscript.exe" 1C77_Utils.vbs Боровая_УпаковкаТаблицИБ.ini
pause