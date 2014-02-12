'option explicit
' Инициализируем необходимые переменные
on error goto 0

' следующие две строки для вывода отладочных сообщений
' эти две строки можно и удалить
Dim DebugFlag 'обязательно глобальная переменная
' DebugFlag = True 'Разрешаю вывод отладочных сообщений

  Dim wshShell
  Dim fso 'as FileSystemObject
  
  Dim LogFile 'as File
  Dim sLogFile 'as string
  Dim ResDict 'as Dictionary 
  'Dim ConfDict 'as Dictionary 
  Dim Status1CLogFileDict 'as Dictionary 
  Dim BatchUploadDict 'as Dictionary 

'      Echo("WScript.ScriptFullName = " + WScript.ScriptFullName)
Wscript.Quit( main() )

'********************************************************************
' Возвращает 1 при успехе, 0 - при неудаче
Function main( )
    main = 0
    
  'Make sure the host is csript, if not then abort
  VerifyHostIsCscript()
  
' проверить версию Windows Script Host
  if CDbl(replace(WScript.Version,".",","))<5.6 then
    Echo "Для работы сценария требуется Windows Script Host версии 5.6 и выше !"
    Exit Function
  end if  

' Инициализация сценария
if not Init() then
  Exit Function
end if

'Exit Function

    InfoBasePath = ResDict.Item(LCase("InfoBasePath")) ' 
    Debug "InfoBasePath", InfoBasePath

    InfoBasesAdminName = CStr(ResDict.Item(LCase("InfoBasesAdminName"))) ' "Администратор1" 'Имя администратора ИБ
    InfoBasesAdminPass = CStr(ResDict.Item(LCase("InfoBasesAdminPass"))) ' "" 'Пароль администратора ИБ

    'Config1CIniFilePath = ResDict.Item(LCase("Config1CIniFilePath")) ' пакетный файл 1С для конфигуратора
    'Log1CPath = ResDict.Item(LCase("1CLogFile")) ' лог конфигуратора из пакетного файла 1С для конфигуратора
    'Debug "Log1CPath", Log1CPath

		    'ServerName = ResDict.Item(LCase("ServerName")) ' "AS-MSK-A6122" 'Имя сервера БД
		    'KlasterPortNumber = ResDict.Item(LCase("KlasterPortNumber")) ' 1541 'Номер пора кластера
		    'InfoBaseName = ResDict.Item(LCase("InfoBaseName")) ' "IMOUT_User_AAyuhanov01" 'Имя ИБ

			'sFullServerName = ServerName
			'if "" <> CStr(KlasterPortNumber) then
			'	sFullServerName = ServerName + ":" + CStr(KlasterPortNumber)
			'end if

		'FilePath = ResDict.Item(LCase("FilePath")) ' "\\AS-MSK-A6122\Share\Admin1C\confupdate.vbs" 'Путь к текущему файлу
    NetFile = ResDict.Item(LCase("NetFile")) ' "\\AS-MSK-A6122\Share\Admin1C\confupdate_base.txt" 'Путь к log-файлу в сети - используется только для NeedCopyFiles = True

    'Folder = ResDict.Item(LCase("Folder")) ' "\\AS-MSK-A6122\Share\Admin1C\" 'Каталог для выгрузки базы

    CountDB = CInt(ResDict.Item(LCase("CountDB"))) ' 7 'За сколько дней хранить копии
    Prefix = ResDict.Item(LCase("Prefix")) ' "base" 'Префикс файла выгрузки
    Out = ResDict.Item(LCase(LCase("OwnLogFile"))) ' "\\AS-MSK-A6122\Share\Admin1C\confupdate.txt" 'Путь к log-файлу
    sLogFile = Out
    Debug "Out", Out

    NeedTestIB = UCase(ResDict.Item(LCase("NeedTestIB"))) = "TRUE" ' False ' Необходимость тестирования базы
				'NeedUpdateFromStorage = UCase(ResDict.Item(LCase("NeedUpdateFromStorage"))) = "TRUE" ' Необходимость обновления конфигурации из хранилища конфигурации
			    'NeedDumpIB = UCase(ResDict.Item(LCase("NeedDumpIB"))) = "TRUE" ' True ' Необходимость выгрузки базы
			    'NeedCopyFiles = UCase(ResDict.Item(LCase("NeedCopyFiles"))) = "TRUE" ' True ' Необходимость выгрузки базы
			    'NeedTestIB = UCase(ResDict.Item(LCase("NeedTestIB"))) = "TRUE" ' False ' Необходимость тестирования базы
			    'NeedRestartAgent = UCase(ResDict.Item(LCase("NeedRestartAgent"))) = "TRUE" ' False ' Необходимость рестарта агента сервера
			    'NeedRestoreIB = UCase(ResDict.Item(LCase("NeedRestoreIB"))) = "TRUE" ' Необходимость восстановления конфигурации из файла
			    'NeedRestoreIB83 = UCase(ResDict.Item(LCase("NeedRestoreIB83"))) = "TRUE" ' Необходимость восстановления конфигурации из файла платформой 8.3
			    '    
			    'IBFile = ResDict.Item(LCase("IBFile")) ' "" 'Путь к файлу с выгрузкой базы
			    'LockMessageText = ResDict.Item(LCase("LockMessageText")) ' "Идет регламент. Подождите..." 'Текст сообщения о блокировки подключений к ИБ
			    'LockPermissionCode = ResDict.Item(LCase("LockPermissionCode")) ' "Артур" 'Ключ для запуска заблокированной ИБ
			    'AuthStr = ResDict.Item(LCase("AuthStr")) ' "/WA+" 
			    'TimeSleep = ResDict.Item(LCase("TimeSleep")) ' 10000 '600000 '10 секунд 600 секунд
			    'TimeSleepShort = ResDict.Item(LCase("TimeSleepShort")) ' 2000 '60000 '2 секунд 60 секунд
			    'Cfg = ResDict.Item(LCase("Cfg")) ' "" 'Путь к файлу с измененной конфигурацией
			    'InfoCfgFile = ResDict.Item(LCase("InfoCfgFile")) ' "" 'Информация о файле обновления конфигурации
			    'v8exe = ResDict.Item(LCase("v8exe")) ' "C:\Program Files (x86)\1cv82\8.2.18.96\bin\1cv8.exe" 'Путь к исполняемому файлу 1С:Предприятия 8.2
				'v83exe = ResDict.Item(LCase("v83exe"))
				'	'rem NewPass = "" 'Новый пароль администратора, обновляющего ИБ

    v7exe = ResDict.Item(LCase("v7exe")) 

    OpenLogFile

    TimeBeginLock = Now ' Время начала блокировки ИБ
    TimeEndLock = DateAdd("h", 2, TimeBeginLock) ' Время окончания блокировки ИБ

	Echo(CStr(Now) + " НАЧАЛО ТЕСТИРОВАНИЯ КОНФИГУРАЦИИ 1С 7.7")
    Echo(CStr(Now) + " Путь к базе " + InfoBasePath)

        'Покажем свободное место на диске с исполняемым файлом 1С
	Echo(CStr(Now) + " " + ShowFreeSpace(v7exe))

	Echo(CStr(Now) + " " + ShowFreeSpace(InfoBasePath))
        'Покажем свободное место на диске с архивами
        'Echo(CStr(Now) + " " + ShowFreeSpace(Folder))

    sUserLoginPass = " /N" + InfoBasesAdminName 
    if InfoBasesAdminPass <> "" then
	    sUserLoginPass = sUserLoginPass + " /P" + InfoBasesAdminPass + " "
	end if

        'If FSO.FolderExists(Folder) = False Then
        '    FSO.CreateFolder Folder
        'End if

	sTempBatchUploadFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()
	sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()

    Fill1CBatchUploadDict
	Create1C77BatchUploadFile sTempBatchUploadFile, sTempFile

        If NeedTestIB = True Then
            Echo(CStr(Now) + " тестируем базу ") ' EchoWithOpenAndCloseLog

            LineExe = """" + v7exe + """ config /D""" + InfoBasePath + """"  + sUserLoginPass + " /@" + sTempBatchUploadFile
            'LineExe = """" + v7exe + """ config /D""" + InfoBasePath + """"  + sUserLoginPass + " /@" + Config1CIniFilePath

            		' /IBCheckAndRepair -LogIntegrity -RecalcTotals /Out""" + sTempFile + """ -NoTruncate"
            Echo(CStr(Now) + " ком.строка: " + LineExe) ' EchoWithOpenAndCloseLog

            wshShell.Run LineExe, 5, True

			success = Show1C77ConfigLog(sTempFile, " ОШИБКА: 1С вернула ошибку при выполнении тестирования и исправления") 'Log1CPath
        End if
        
        'OpenLogFile


    Echo(CStr(Now) + " ЗАВЕРШЕНИЕ ТЕСТИРОВАНИЯ КОНФИГУРАЦИИ 1С 7.7")
    'WriteLogIntoIBEventLog sFullServerName, InfoBaseName, sLogFile

    If NeedCopyFiles = True Then 
        If fso.FileExists(NetFile) Then
            fso.DeleteFile(NetFile)
        End If
        fso.MoveFile Out, NetFile
    End if

    If NeedDumpIB = True Then 
        CALL DelOldFiles(Folder, CountDB)
    End if

    if not success then
    	main = 1
    end if
End Function

Function Create1C77BatchUploadFile(sFileName, sTempLogFileName)
'on error resume next
	Debug "Create1C77BatchUploadFile sFileName", sFileName
    ForWriting = 2
    Set File = fso.OpenTextFile(sFileName, ForWriting, True)

   	File.WriteLine "[General]"
   	File.WriteLine "Output=" + sTempLogFileName
   	File.WriteLine "Quit=1"
   	File.WriteLine "CheckAndRepair=1"
   	File.WriteLine ""

   	File.WriteLine "[CheckAndRepair]"
   	File.WriteLine "Repair=1"
   	File.WriteLine "PhysicalIntegrity=" + BatchUploadDict.Item(LCase("PhysicalIntegrity"))
   	File.WriteLine "Reindex=" + BatchUploadDict.Item(LCase("Reindex"))
   	File.WriteLine "LogicalIntegrity=" + BatchUploadDict.Item(LCase("LogicalIntegrity"))
   	File.WriteLine "RecalcSecondaries=" + BatchUploadDict.Item(LCase("RecalcSecondaries"))
   	File.WriteLine "RecalcTotals=" + BatchUploadDict.Item(LCase("RecalcTotals"))
   	File.WriteLine "Pack=" + BatchUploadDict.Item(LCase("Pack"))

			'//SkipUnresolved=
			'//CreateForUnresolved=
			'//Reconstruct=	

	file.Close
'on error goto 0
End Function

Sub Fill1CBatchUploadDict()
  Dim elem, key, value
  dim keys, items
  keys = BatchUploadDict.Keys()
  items = BatchUploadDict.Items()
  For elem=0 To BatchUploadDict.Count-1
    'name = dict(elem)
    key = keys(elem)
    value = items(elem)
    value = ResDict.Item(LCase(key))
debug "ResDict("+CStr(key)+")", value
    if LCase(value) = "true" then
    	value = "1"
    else
    	value = "0"
    end if
    BatchUploadDict.Remove key
    BatchUploadDict.Add LCase(key), value
debug "BatchUploadDict("+CStr(LCase(key))+")", value
  Next 'elem
End Sub

' получить данные из INI-файла
' ResDict - объект Dictionary, где хранятся пары ключ/значение
Function Show1C77ConfigLog(LogFileName, errorMessage)
    Dim File 'As TextStream

    On Error Resume Next
    Dim ForRead
    ForRead =1
    Set File = fso.OpenTextFile(LogFileName,ForRead)
    if err.Number<>0 then
      Err.Clear()
      echo "log-файл "& LogFileName &" не удалось открыть!"
      Exit Function
    end if
    on error goto 0

    Set ResDict = CreateObject("Scripting.Dictionary")
    Dim s, Matches, Match
    Dim reg 'As RegExp
    Set reg = new RegExp
      reg.Pattern= "^(\d{4})(\d{2})(\d{2});([^;]+);[^;]*;C;Doctor;(\w+);(\d);([^;]*);;" ' "^\s*([^=]+)\s*=\s*([^;']+)[;']?"
					' нормальная строка вида '20130919;16:51:04;;C;Doctor;dctPhInt;1;;;
					'или ошибка '20130919;21:13:43;Администратор;C;Doctor;dctErr;5;Создана таблица - 1SBLOB;;
      reg.IgnoreCase = True

    Dim elem, index

'DebugDict Status1CLogFileDict
	success = true
	normalFinish = false

    Do While File.AtEndOfStream <> True
      s = File.ReadLine
    ' если не строка-комментарий  
      if not RegExpTest("^\s*[;']",s) then
    '  For index=0 To IniDict.Count-1
    '    reg.Pattern="\s*"+elem(index)+"\s*=\s*(.+)"
    ' выделить ключ и значение в Ini-файле, убрав возможные комментарии
        Set Matches = reg.Execute(s)
        if Matches.Count>0 then
   
	        Dim sDateTime, sConfigActionKey, sConfigActionFull, bConfigActionRes, sMessage

			sDateTime = Matches(0).SubMatches(2) + "." + Matches(0).SubMatches(1) + "." + Matches(0).SubMatches(0) + " " + Matches(0).SubMatches(3)
			sConfigActionKey = Matches(0).SubMatches(4)

			if sConfigActionKey = "dctTREnd" then ' 1C может вылететь при тестировании или сеанс может кто-то завершить принудительно, нужно проверять нормальное завершение процесса
				normalFinish = true
			end if
			
			sConfigActionFull = Status1CLogFileDict.Item(sConfigActionKey)

	        bConfigActionRes = Matches(0).SubMatches(5) = 1
			
	        sMessage = sDateTime + " " + sConfigActionFull
			if IsEmpty(sConfigActionFull) then
	        	sMessage = sDateTime + " " + Matches(0).SubMatches(6)
			end if

	        if bConfigActionRes then
	        	sMessage = sMessage + " : выполнено успешно"
	        else
				success = false
	        	sMessage = sMessage + " : обнаружена ошибка!"
	        end if
	            
			Echo sMessage

			ResDict.Add sConfigActionKey, sMessage


'Debug "<sConfigActionKey = sMessage>", sConfigActionKey + " = [" + sMessage + "]"

		else ' if Matches.Count>0 then
			success = false ' на всякий случай, для анализа неизвестной строки, полученной от Конфигуратора
			Echo sMessage
        end if
      end if
    Loop
    File.Close()

			'if haveProblem = true or not normalFinish then
	if not success  or not normalFinish then
		Echo(CStr(Now) + errorMessage) ' EchoWithOpenAndCloseLog '" Ошибка при обновлении конфигурации из хранилища")
	end if

    if ResDict.Count=0 then
      echo "Не удалось прочесть данные из log-файла " & LogFileName
      Show1C77ConfigLog = false
    else  
      Show1C77ConfigLog = success
    end if
End Function 'GetDataFrom1C77LogFile

Function Show1CConfigLog(sTempFile, errorMessage)
	Set configLogFile = fso.OpenTextFile(sTempFile, 1)

	success = true
	Do While configLogFile.AtEndOfStream <> True
		errorString = configLogFile.ReadLine
		Echo errorString
		errorPos = InStr(1, lCase(errorString), "ошибка", 1)
		If errorPos > 0 Then
			success = false
		end if
	Loop
	if not success then
		Echo(CStr(Now) + errorMessage) ' EchoWithOpenAndCloseLog '" Ошибка при обновлении конфигурации из хранилища")
	end if
	configLogFile.Close()

	Show1C77ConfigLog = success
End Function

Sub WriteLogIntoIBEventLog(sFullServerName, InfoBaseName, sLogFile)
		'Sub WriteLogIntoIBEventLog(ServerName, KlasterPortNumber, InfoBaseName, sLogFile)
    Echo(CStr(Now) + " Сохранение лога в журнал регистрации ИБ")
    Set ComConnector = CreateObject("v82.COMConnector")
        'Set connection = ComConnector.Connect("Srvr=" + ServerName + ":" + CStr(KlasterPortNumber) + ";Ref=" + InfoBaseName + ";Usr=" + InfoBasesAdminName + ";Pwd=" + InfoBasesAdminPass)
    Set connection = ComConnector.Connect("Srvr=" + sFullServerName + ";Ref=" + InfoBaseName)

    Echo(CStr(Now) + " ЗАВЕРШЕНИЕ ОБНОВЛЕНИЯ КОНФИГУРАЦИИ")

    'LogFile.Close()
    'LogFile = ""

    Set f = fso.OpenTextFile(sLogFile, 1, False, -2) 'Out
    Text = f.ReadAll

    'Запишем всю информацию из log-файла в журнал регистрации
    connection.WriteLogEvent "Регламентное обновление ИБ", connection.EventLogLevel.Information,,, Text

    connection = Null
    ComConnector = Null
    f = Null
End Sub

Sub TerminateProcess(strProcessName)
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & strProcessName & "'")
    For Each objProcess in colProcess
        objProcess.Terminate()
        Echo(CStr(Now) + " " + CStr(objProcess.Name) + " Завершение процесса агента сервера 1С Предприятия")
    Next
End Sub

' Инициализация сценария
Function Init( )
      Init = false
        
      set wshShell = wScript.createObject("wScript.shell")
      Set fso = CreateObject("Scripting.FileSystemObject") 
      
    ' задать имя ini-файла
      Dim IniFileName

      Dim intOpMode
        intOpMode = intParseCmdLine(IniFileName)

			' всегда один ini-файл в каталоге программы
			'  IniFileName = Replace(LCase(WScript.ScriptFullName),".vbs",".ini")
    Debug "IniFileName", IniFileName

      if not GetDataFromIniFile(IniFileName, ResDict) then
        Exit Function
      end if

    On Error Resume Next
      Dim sDebugFlag
      sDebugFlag = ResDict.Item(LCase("DebugFlag"))
      Debug "sDebugFlag",sDebugFlag
      if sDebugFlag<>"" then
        DebugFlag = CBool(sDebugFlag)
      end if
      Debug "DebugFlag",DebugFlag
    On Error Goto 0

        '' получить лог-файл
        '  LogFile = Null 'не выводить в лог-файл, если не задан путь к нему
        '  Dim sLogFile
        '  sLogFile = ResDict.Item(LCase("LogFile"))
        '  if sLogFile<>"" then
        '    If (NOT blnOpenFile(sLogFile, LogFile)) Then
        '      Call Wscript.Echo ("Не могу открыть лог-файл <"+sLogFile+"> .")
        '      Exit Function
        '    End If
        '  End If    

  CreateStatus1CLogFileDict
  Create1CBatchUploadDict

  Init = true
End Function 'Init      

Sub CreateStatus1CLogFileDict()
    Set Status1CLogFileDict = CreateObject("Scripting.Dictionary")
	
    Status1CLogFileDict.Add "dctTRBeg", "Начало тестирования и исправления"
    Status1CLogFileDict.Add "dctPhInt", "Контроль физической целостности"
    Status1CLogFileDict.Add "dctReind", "Реиндексация таблиц ИБ"
    Status1CLogFileDict.Add "dctLgInt", "Контроль логической целостности"
    Status1CLogFileDict.Add "dctRcST", "Пересчет служебных данных"
    Status1CLogFileDict.Add "dctRcT", "Пересчет итогов"
    Status1CLogFileDict.Add "dctPck", "Упаковка таблиц ИБ"
    Status1CLogFileDict.Add "dctTREnd", "Окончание тестирования и исправления ИБ"

		'20130919;16:51:04;;C;Doctor;dctTRBeg;1;;;
		'Начало тестирования и исправления
		'20130919;16:51:04;;C;Doctor;dctPhInt;1;;;
		'Контроль физической целостности
		'20130919;16:51:04;;C;Doctor;dctReind;1;;;
		'Реиндексация таблиц ИБ
		'20130919;16:51:04;;C;Doctor;dctLgInt;1;;;
		'Контроль логической целостности
		'20130919;16:51:04;;C;Doctor;dctRcST;1;;;
		'Пересчет служебных данных
		'20130919;16:51:04;;C;Doctor;dctRcT;1;;;
		'Пересчет итогов
		'20130919;16:51:04;;C;Doctor;dctPck;1;;;
		'Упаковка таблиц ИБ
		'20130919;16:51:04;;C;Doctor;dctTREnd;1;;;
		'Окончание тестирования и исправления ИБ
End sub

Sub Create1CBatchUploadDict()
    Set BatchUploadDict = CreateObject("Scripting.Dictionary")

	BatchUploadDict.Add "PhysicalIntegrity", 0 ' Контроль физической целостности
	BatchUploadDict.Add "Reindex", 0 ' Реиндексация таблиц ИБ
	BatchUploadDict.Add "LogicalIntegrity", 0 ' Контроль логической целостности
	BatchUploadDict.Add "RecalcSecondaries", 0 ' Пересчет служебных данных
	BatchUploadDict.Add "RecalcTotals", 0 ' Пересчет итогов
	BatchUploadDict.Add "Pack", 0 ' Упаковка таблиц ИБ
End sub

Function GetFormatDay()
    iDay = Day(Now)
    mDay = CStr(Day(Now))
    iMonth = Month(Now)
    mMonth = CStr(Month(Now))
    mYear = CStr(Year(Now))

    nCDay = "_" + mYear + "_"
    If iMonth < 10 Then
       nCDay = nCDay + "0"
    End If
    nCDay = nCDay + mMonth + "_"
    If iDay < 10 Then
       nCDay = nCDay + "0"
    End If
    nCDay = nCDay + mDay
	
	GetFormatDay = nCDay
End Function

' Функция для определения свободного места на диске
Function ShowFreeSpace(drvPath)
  Dim d, s
  on error Resume next
  Set d = fso.GetDrive(fso.GetDriveName(drvPath))
  s = "Drive " & UCase(drvPath) & " - " 
  s = s & d.VolumeName  & " "
  s = s & "Free Space: " & FormatNumber(d.FreeSpace/1024/1024, 0) 
  s = s & " Mbytes"
  on error goto 0
  ShowFreeSpace = s
End Function

' Скрипт для затирания устаревших файлов: 
' Удаляет только файлы у которых сходятся префиксы
Sub DelOldFiles(Folder_Name, Stack_Depth)
    Set folder = fso.GetFolder(Folder_Name)
    Set files = folder.Files
    For Each f in files
        fdate = f.DateCreated
        fPrefix = Left(f.Name,Len(Prefix))
        If ((Date - fdate) > Stack_Depth) And fPrefix = Prefix Then
            f.Delete
        End If
    Next
End Sub

' получить данные из INI-файла
' ResDict - объект Dictionary, где хранятся пары ключ/значение
Function GetDataFromIniFile(ByVal IniFileName, ByRef ResDict)
      GetDataFromIniFile = false
  
    ' далее автоматически
    Dim IniFile 'As TextStream

    On Error Resume Next
    Dim ForRead
    ForRead =1
    Set IniFile = fso.OpenTextFile(IniFileName,ForRead)
    if err.Number<>0 then
      Err.Clear()
      echo "Ini-файл "& IniFileName &" не удалось открыть!"
      Exit Function
    end if
    on error goto 0

    Set ResDict = CreateObject("Scripting.Dictionary")
    Dim s, Matches, Match
    Dim reg 'As RegExp
    Set reg = new RegExp
      reg.Pattern="^\s*([^=]+)\s*=\s*([^;']+)[;']?"
      reg.IgnoreCase = True

    Dim elem, index

    Do While IniFile.AtEndOfStream <> True
      s = IniFile.ReadLine
    ' если не строка-комментарий  
      if not RegExpTest("^\s*[;']",s) then
    '  For index=0 To IniDict.Count-1
    '    reg.Pattern="\s*"+elem(index)+"\s*=\s*(.+)"
    ' выделить ключ и значение в Ini-файле, убрав возможные комментарии
        Set Matches = reg.Execute(s)
        if Matches.Count>0 then
   
        Dim lkey, lvalue
		' добавить новую пару, исключив из значения табуляцию и левые(и правые) пробелы    
					'ResDict.Add elem(index),Trim(replace(Matches(0).SubMatches(0),vbTab," "))
            lkey = LCase(Trim(replace(Matches(0).SubMatches(0),vbTab," ")))
            lvalue = replace(Matches(0).SubMatches(1), vbTab, " ")
            lvalue = Trim(replace(lvalue, chr(34), "")) 'убираю кавычки
            
            ResDict.Add lkey, lvalue
					'ResDict.Add LCase(Trim(replace(Matches(0).SubMatches(0),vbTab," "))),Trim(replace(Matches(0).SubMatches(1),vbTab," "))

Debug "lkey=lvalue", lkey + " = [" + lvalue + "]"
        end if
      end if
    Loop
    IniFile.Close()

    if ResDict.Count=0 then
      echo "Не удалось прочесть данные из Ini-файла " & IniFileName
      GetDataFromIniFile = false
    else  
      GetDataFromIniFile = true
    end if
End Function 'GetDataFromIniFile


' проверить на соответствие шаблону
' регистр символов не важен
  Dim regExTest               ' Create variable.
Function RegExpTest(ByVal patrn, ByVal strng)
  if IsEmpty(regExTest) then
    Set regExTest = New RegExp         ' Create regular expression.
  end if
  regExTest.Pattern = patrn         ' Set pattern.
  regExTest.IgnoreCase = true      ' disable case sensitivity.
  RegExpTest = regExTest.Test(strng)      ' Execute the search test.
'  regEx = null
End Function

Function OpenLogFile()
	Debug "sLogFile", sLogFile
    Set LogFile = fso.OpenTextFile(sLogFile, 8, True)
    set OpenLogFile = LogFile
End Function

Sub Echo(text)
  WScript.Echo(text)
on error resume next
  If IsObject(LogFile) then        'LogFile should be a file object
    LogFile.WriteLine text
  end if
on error goto 0
End Sub'Echo

Sub EchoWithOpenAndCloseLog(text)
    OpenLogFile

    Echo(text)    

    LogFile.Close()    
    LogFile = ""
End Sub'Echo

Sub Debug(ByVal title, ByVal msg)
'exit sub
on error resume next
  DebugFlag = DebugFlag
  if err.Number<>0 then
    err.Clear()
    on error goto 0
    Exit Sub
  end if
  if DebugFlag then
    if not (IsEmpty(msg) or IsNull(msg)) then
      msg = CStr(msg)
    end if
    if not (IsEmpty(title) or IsNull(title)) then
      title = CStr(title)
    end if
    If msg="" Then
      Echo(title)
    else
      Echo(title+" - <"+msg+">")
    End If
  End If
on error goto 0
End Sub'Debug

Sub DebugDict(dict)
  Dim elem, key, value
  dim keys, items
  keys = dict.Keys()
  items = dict.Items()
  For elem=0 To dict.Count-1
    'name = dict(elem)
    key = keys(elem)
    value = items(elem)
debug "dict("+CStr(key)+")", value
  Next 'elem
End Sub

Private Function intParseCmdLine( ByRef strFileName)

	Dim strFlag 'intParseCmdLine

'    ON ERROR RESUME NEXT
    If Wscript.Arguments.Count > 0 Then
        strFlag = Wscript.arguments.Item(0)
    End If

    If IsEmpty(strFlag) Then                'No arguments have been received
        ShowUsage 'intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

        'Check if the user is asking for help or is just confused
    If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") _
        OR (strFlag="h") Then
        ShowUsage 'intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If
    intParseCmdLine = 0 'CONST_LIST

    strFilename = strFlag
End Function

Sub ShowUsage ()

    Wscript.Echo ""
'    Wscript.Echo "Копирует файл на диск A:. Там же создается копия файла."
    Wscript.Echo "Выполняет административные действия с базой 1С 8.2"
    Wscript.Echo ""
    Wscript.Echo "ПАРАМЕТРЫ ВЫЗОВА:"
    Wscript.Echo "  "+ WScript.ScriptName +" [файл-настроек | /? | /h]"
    Wscript.Echo ""
    Wscript.Echo "ПРИМЕР:"
    Wscript.Echo "1. cscript "+ WScript.ScriptName +" Файл.ini"
    Wscript.Echo "2. cscript "+ WScript.ScriptName
    Wscript.Echo "   Показывает этот экран."

End Sub

'********************************************************************
'* 
'* Function blnOpenFile
'*
'* Purpose: Opens a file.
'*
'* Input:   strFileName         A string with the name of the file.
'*
'* Output:  Sets objOpenFile to a FileSystemObject and setis it to 
'*            Nothing upon Failure.
'* 
'********************************************************************
Private Function blnOpenFile(ByVal strFileName, ByRef objOpenFile)

    ON ERROR RESUME NEXT

    If IsEmpty(strFileName) OR strFileName = "" Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If

    'fso.DeleteFile(strFileName)
    'Open the file for output
    Set objOpenFile = fso.CreateTextFile(strFileName, True)
    If blnErrorOccurred("Невозможно создать") Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If
    blnOpenFile = True

End Function

'********************************************************************
'*
'* Sub      VerifyHostIsCscript()
'*
'* Purpose: Determines which program is used to run this script.
'*
'* Input:   None
'*
'* Output:  If host is not cscript, then an error message is printed 
'*          and the script is aborted.
'*
'********************************************************************
Sub VerifyHostIsCscript()

    ON ERROR RESUME NEXT

    'Define constants
    CONST CONST_ERROR                   = 0
    CONST CONST_WSCRIPT                 = 1
    CONST CONST_CSCRIPT                 = 2
    
    Dim strFullName, strCommand, i, j, intStatus

    strFullName = WScript.FullName

    If Err.Number then
        Call Echo( "Error 0x" & CStr(Hex(Err.Number)) & " occurred." )
        If Err.Description <> "" Then
            Call Echo( "Error description: " & Err.Description & "." )
        End If
        intStatus =  CONST_ERROR
    End If

    i = InStr(1, strFullName, ".exe", 1)
    If i = 0 Then
        intStatus =  CONST_ERROR
    Else
        j = InStrRev(strFullName, "\", i, 1)
        If j = 0 Then
            intStatus =  CONST_ERROR
        Else
            strCommand = Mid(strFullName, j+1, i-j-1)
            Select Case LCase(strCommand)
                Case "cscript"
                    intStatus = CONST_CSCRIPT
                Case "wscript"
                    intStatus = CONST_WSCRIPT
                Case Else       'should never happen
                    Call Echo( "An unexpected program was used to " _
                                       & "run this script." )
                    Call Echo( "Only CScript.Exe or WScript.Exe can " _
                                       & "be used to run this script." )
                    intStatus = CONST_ERROR
                End Select
        End If
    End If

    If intStatus <> CONST_CSCRIPT Then
        Call Echo( "Please run this script using CScript." & vbCRLF & _
             "This can be achieved by" & vbCRLF & _
             "1. Using ""CScript SystemAccount.vbs arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""SystemAccount.vbs arguments"" for Windows NT/2000/XP." )
        WScript.Quit(0)
    End If
End Sub 'VerifyHostIsCscript
