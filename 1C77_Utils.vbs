'option explicit
' �������������� ����������� ����������
on error goto 0

' ��������� ��� ������ ��� ������ ���������� ���������
' ��� ��� ������ ����� � �������
Dim DebugFlag '����������� ���������� ����������
' DebugFlag = True '�������� ����� ���������� ���������

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
' ���������� 1 ��� ������, 0 - ��� �������
Function main( )
    main = 0
    
  'Make sure the host is csript, if not then abort
  VerifyHostIsCscript()
  
' ��������� ������ Windows Script Host
  if CDbl(replace(WScript.Version,".",","))<5.6 then
    Echo "��� ������ �������� ��������� Windows Script Host ������ 5.6 � ���� !"
    Exit Function
  end if  

' ������������� ��������
if not Init() then
  Exit Function
end if

'Exit Function

    InfoBasePath = ResDict.Item(LCase("InfoBasePath")) ' 
    Debug "InfoBasePath", InfoBasePath

    InfoBasesAdminName = CStr(ResDict.Item(LCase("InfoBasesAdminName"))) ' "�������������1" '��� �������������� ��
    InfoBasesAdminPass = CStr(ResDict.Item(LCase("InfoBasesAdminPass"))) ' "" '������ �������������� ��

    'Config1CIniFilePath = ResDict.Item(LCase("Config1CIniFilePath")) ' �������� ���� 1� ��� �������������
    'Log1CPath = ResDict.Item(LCase("1CLogFile")) ' ��� ������������� �� ��������� ����� 1� ��� �������������
    'Debug "Log1CPath", Log1CPath

		    'ServerName = ResDict.Item(LCase("ServerName")) ' "AS-MSK-A6122" '��� ������� ��
		    'KlasterPortNumber = ResDict.Item(LCase("KlasterPortNumber")) ' 1541 '����� ���� ��������
		    'InfoBaseName = ResDict.Item(LCase("InfoBaseName")) ' "IMOUT_User_AAyuhanov01" '��� ��

			'sFullServerName = ServerName
			'if "" <> CStr(KlasterPortNumber) then
			'	sFullServerName = ServerName + ":" + CStr(KlasterPortNumber)
			'end if

		'FilePath = ResDict.Item(LCase("FilePath")) ' "\\AS-MSK-A6122\Share\Admin1C\confupdate.vbs" '���� � �������� �����
    NetFile = ResDict.Item(LCase("NetFile")) ' "\\AS-MSK-A6122\Share\Admin1C\confupdate_base.txt" '���� � log-����� � ���� - ������������ ������ ��� NeedCopyFiles = True

    'Folder = ResDict.Item(LCase("Folder")) ' "\\AS-MSK-A6122\Share\Admin1C\" '������� ��� �������� ����

    CountDB = CInt(ResDict.Item(LCase("CountDB"))) ' 7 '�� ������� ���� ������� �����
    Prefix = ResDict.Item(LCase("Prefix")) ' "base" '������� ����� ��������
    Out = ResDict.Item(LCase(LCase("OwnLogFile"))) ' "\\AS-MSK-A6122\Share\Admin1C\confupdate.txt" '���� � log-�����
    sLogFile = Out
    Debug "Out", Out

    NeedTestIB = UCase(ResDict.Item(LCase("NeedTestIB"))) = "TRUE" ' False ' ������������� ������������ ����
				'NeedUpdateFromStorage = UCase(ResDict.Item(LCase("NeedUpdateFromStorage"))) = "TRUE" ' ������������� ���������� ������������ �� ��������� ������������
			    'NeedDumpIB = UCase(ResDict.Item(LCase("NeedDumpIB"))) = "TRUE" ' True ' ������������� �������� ����
			    'NeedCopyFiles = UCase(ResDict.Item(LCase("NeedCopyFiles"))) = "TRUE" ' True ' ������������� �������� ����
			    'NeedTestIB = UCase(ResDict.Item(LCase("NeedTestIB"))) = "TRUE" ' False ' ������������� ������������ ����
			    'NeedRestartAgent = UCase(ResDict.Item(LCase("NeedRestartAgent"))) = "TRUE" ' False ' ������������� �������� ������ �������
			    'NeedRestoreIB = UCase(ResDict.Item(LCase("NeedRestoreIB"))) = "TRUE" ' ������������� �������������� ������������ �� �����
			    'NeedRestoreIB83 = UCase(ResDict.Item(LCase("NeedRestoreIB83"))) = "TRUE" ' ������������� �������������� ������������ �� ����� ���������� 8.3
			    '    
			    'IBFile = ResDict.Item(LCase("IBFile")) ' "" '���� � ����� � ��������� ����
			    'LockMessageText = ResDict.Item(LCase("LockMessageText")) ' "���� ���������. ���������..." '����� ��������� � ���������� ����������� � ��
			    'LockPermissionCode = ResDict.Item(LCase("LockPermissionCode")) ' "�����" '���� ��� ������� ��������������� ��
			    'AuthStr = ResDict.Item(LCase("AuthStr")) ' "/WA+" 
			    'TimeSleep = ResDict.Item(LCase("TimeSleep")) ' 10000 '600000 '10 ������ 600 ������
			    'TimeSleepShort = ResDict.Item(LCase("TimeSleepShort")) ' 2000 '60000 '2 ������ 60 ������
			    'Cfg = ResDict.Item(LCase("Cfg")) ' "" '���� � ����� � ���������� �������������
			    'InfoCfgFile = ResDict.Item(LCase("InfoCfgFile")) ' "" '���������� � ����� ���������� ������������
			    'v8exe = ResDict.Item(LCase("v8exe")) ' "C:\Program Files (x86)\1cv82\8.2.18.96\bin\1cv8.exe" '���� � ������������ ����� 1�:����������� 8.2
				'v83exe = ResDict.Item(LCase("v83exe"))
				'	'rem NewPass = "" '����� ������ ��������������, ������������ ��

    v7exe = ResDict.Item(LCase("v7exe")) 

    OpenLogFile

    TimeBeginLock = Now ' ����� ������ ���������� ��
    TimeEndLock = DateAdd("h", 2, TimeBeginLock) ' ����� ��������� ���������� ��

	Echo(CStr(Now) + " ������ ������������ ������������ 1� 7.7")
    Echo(CStr(Now) + " ���� � ���� " + InfoBasePath)

        '������� ��������� ����� �� ����� � ����������� ������ 1�
	Echo(CStr(Now) + " " + ShowFreeSpace(v7exe))

	Echo(CStr(Now) + " " + ShowFreeSpace(InfoBasePath))
        '������� ��������� ����� �� ����� � ��������
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
            Echo(CStr(Now) + " ��������� ���� ") ' EchoWithOpenAndCloseLog

            LineExe = """" + v7exe + """ config /D""" + InfoBasePath + """"  + sUserLoginPass + " /@" + sTempBatchUploadFile
            'LineExe = """" + v7exe + """ config /D""" + InfoBasePath + """"  + sUserLoginPass + " /@" + Config1CIniFilePath

            		' /IBCheckAndRepair -LogIntegrity -RecalcTotals /Out""" + sTempFile + """ -NoTruncate"
            Echo(CStr(Now) + " ���.������: " + LineExe) ' EchoWithOpenAndCloseLog

            wshShell.Run LineExe, 5, True

			success = Show1C77ConfigLog(sTempFile, " ������: 1� ������� ������ ��� ���������� ������������ � �����������") 'Log1CPath
        End if
        
        'OpenLogFile


    Echo(CStr(Now) + " ���������� ������������ ������������ 1� 7.7")
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

' �������� ������ �� INI-�����
' ResDict - ������ Dictionary, ��� �������� ���� ����/��������
Function Show1C77ConfigLog(LogFileName, errorMessage)
    Dim File 'As TextStream

    On Error Resume Next
    Dim ForRead
    ForRead =1
    Set File = fso.OpenTextFile(LogFileName,ForRead)
    if err.Number<>0 then
      Err.Clear()
      echo "log-���� "& LogFileName &" �� ������� �������!"
      Exit Function
    end if
    on error goto 0

    Set ResDict = CreateObject("Scripting.Dictionary")
    Dim s, Matches, Match
    Dim reg 'As RegExp
    Set reg = new RegExp
      reg.Pattern= "^(\d{4})(\d{2})(\d{2});([^;]+);[^;]*;C;Doctor;(\w+);(\d);([^;]*);;" ' "^\s*([^=]+)\s*=\s*([^;']+)[;']?"
					' ���������� ������ ���� '20130919;16:51:04;;C;Doctor;dctPhInt;1;;;
					'��� ������ '20130919;21:13:43;�������������;C;Doctor;dctErr;5;������� ������� - 1SBLOB;;
      reg.IgnoreCase = True

    Dim elem, index

'DebugDict Status1CLogFileDict
	success = true
	normalFinish = false

    Do While File.AtEndOfStream <> True
      s = File.ReadLine
    ' ���� �� ������-�����������  
      if not RegExpTest("^\s*[;']",s) then
    '  For index=0 To IniDict.Count-1
    '    reg.Pattern="\s*"+elem(index)+"\s*=\s*(.+)"
    ' �������� ���� � �������� � Ini-�����, ����� ��������� �����������
        Set Matches = reg.Execute(s)
        if Matches.Count>0 then
   
	        Dim sDateTime, sConfigActionKey, sConfigActionFull, bConfigActionRes, sMessage

			sDateTime = Matches(0).SubMatches(2) + "." + Matches(0).SubMatches(1) + "." + Matches(0).SubMatches(0) + " " + Matches(0).SubMatches(3)
			sConfigActionKey = Matches(0).SubMatches(4)

			if sConfigActionKey = "dctTREnd" then ' 1C ����� �������� ��� ������������ ��� ����� ����� ���-�� ��������� �������������, ����� ��������� ���������� ���������� ��������
				normalFinish = true
			end if
			
			sConfigActionFull = Status1CLogFileDict.Item(sConfigActionKey)

	        bConfigActionRes = Matches(0).SubMatches(5) = 1
			
	        sMessage = sDateTime + " " + sConfigActionFull
			if IsEmpty(sConfigActionFull) then
	        	sMessage = sDateTime + " " + Matches(0).SubMatches(6)
			end if

	        if bConfigActionRes then
	        	sMessage = sMessage + " : ��������� �������"
	        else
				success = false
	        	sMessage = sMessage + " : ���������� ������!"
	        end if
	            
			Echo sMessage

			ResDict.Add sConfigActionKey, sMessage


'Debug "<sConfigActionKey = sMessage>", sConfigActionKey + " = [" + sMessage + "]"

		else ' if Matches.Count>0 then
			success = false ' �� ������ ������, ��� ������� ����������� ������, ���������� �� �������������
			Echo sMessage
        end if
      end if
    Loop
    File.Close()

			'if haveProblem = true or not normalFinish then
	if not success  or not normalFinish then
		Echo(CStr(Now) + errorMessage) ' EchoWithOpenAndCloseLog '" ������ ��� ���������� ������������ �� ���������")
	end if

    if ResDict.Count=0 then
      echo "�� ������� �������� ������ �� log-����� " & LogFileName
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
		errorPos = InStr(1, lCase(errorString), "������", 1)
		If errorPos > 0 Then
			success = false
		end if
	Loop
	if not success then
		Echo(CStr(Now) + errorMessage) ' EchoWithOpenAndCloseLog '" ������ ��� ���������� ������������ �� ���������")
	end if
	configLogFile.Close()

	Show1C77ConfigLog = success
End Function

Sub WriteLogIntoIBEventLog(sFullServerName, InfoBaseName, sLogFile)
		'Sub WriteLogIntoIBEventLog(ServerName, KlasterPortNumber, InfoBaseName, sLogFile)
    Echo(CStr(Now) + " ���������� ���� � ������ ����������� ��")
    Set ComConnector = CreateObject("v82.COMConnector")
        'Set connection = ComConnector.Connect("Srvr=" + ServerName + ":" + CStr(KlasterPortNumber) + ";Ref=" + InfoBaseName + ";Usr=" + InfoBasesAdminName + ";Pwd=" + InfoBasesAdminPass)
    Set connection = ComConnector.Connect("Srvr=" + sFullServerName + ";Ref=" + InfoBaseName)

    Echo(CStr(Now) + " ���������� ���������� ������������")

    'LogFile.Close()
    'LogFile = ""

    Set f = fso.OpenTextFile(sLogFile, 1, False, -2) 'Out
    Text = f.ReadAll

    '������� ��� ���������� �� log-����� � ������ �����������
    connection.WriteLogEvent "������������ ���������� ��", connection.EventLogLevel.Information,,, Text

    connection = Null
    ComConnector = Null
    f = Null
End Sub

Sub TerminateProcess(strProcessName)
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & strProcessName & "'")
    For Each objProcess in colProcess
        objProcess.Terminate()
        Echo(CStr(Now) + " " + CStr(objProcess.Name) + " ���������� �������� ������ ������� 1� �����������")
    Next
End Sub

' ������������� ��������
Function Init( )
      Init = false
        
      set wshShell = wScript.createObject("wScript.shell")
      Set fso = CreateObject("Scripting.FileSystemObject") 
      
    ' ������ ��� ini-�����
      Dim IniFileName

      Dim intOpMode
        intOpMode = intParseCmdLine(IniFileName)

			' ������ ���� ini-���� � �������� ���������
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

        '' �������� ���-����
        '  LogFile = Null '�� �������� � ���-����, ���� �� ����� ���� � ����
        '  Dim sLogFile
        '  sLogFile = ResDict.Item(LCase("LogFile"))
        '  if sLogFile<>"" then
        '    If (NOT blnOpenFile(sLogFile, LogFile)) Then
        '      Call Wscript.Echo ("�� ���� ������� ���-���� <"+sLogFile+"> .")
        '      Exit Function
        '    End If
        '  End If    

  CreateStatus1CLogFileDict
  Create1CBatchUploadDict

  Init = true
End Function 'Init      

Sub CreateStatus1CLogFileDict()
    Set Status1CLogFileDict = CreateObject("Scripting.Dictionary")
	
    Status1CLogFileDict.Add "dctTRBeg", "������ ������������ � �����������"
    Status1CLogFileDict.Add "dctPhInt", "�������� ���������� �����������"
    Status1CLogFileDict.Add "dctReind", "������������ ������ ��"
    Status1CLogFileDict.Add "dctLgInt", "�������� ���������� �����������"
    Status1CLogFileDict.Add "dctRcST", "�������� ��������� ������"
    Status1CLogFileDict.Add "dctRcT", "�������� ������"
    Status1CLogFileDict.Add "dctPck", "�������� ������ ��"
    Status1CLogFileDict.Add "dctTREnd", "��������� ������������ � ����������� ��"

		'20130919;16:51:04;;C;Doctor;dctTRBeg;1;;;
		'������ ������������ � �����������
		'20130919;16:51:04;;C;Doctor;dctPhInt;1;;;
		'�������� ���������� �����������
		'20130919;16:51:04;;C;Doctor;dctReind;1;;;
		'������������ ������ ��
		'20130919;16:51:04;;C;Doctor;dctLgInt;1;;;
		'�������� ���������� �����������
		'20130919;16:51:04;;C;Doctor;dctRcST;1;;;
		'�������� ��������� ������
		'20130919;16:51:04;;C;Doctor;dctRcT;1;;;
		'�������� ������
		'20130919;16:51:04;;C;Doctor;dctPck;1;;;
		'�������� ������ ��
		'20130919;16:51:04;;C;Doctor;dctTREnd;1;;;
		'��������� ������������ � ����������� ��
End sub

Sub Create1CBatchUploadDict()
    Set BatchUploadDict = CreateObject("Scripting.Dictionary")

	BatchUploadDict.Add "PhysicalIntegrity", 0 ' �������� ���������� �����������
	BatchUploadDict.Add "Reindex", 0 ' ������������ ������ ��
	BatchUploadDict.Add "LogicalIntegrity", 0 ' �������� ���������� �����������
	BatchUploadDict.Add "RecalcSecondaries", 0 ' �������� ��������� ������
	BatchUploadDict.Add "RecalcTotals", 0 ' �������� ������
	BatchUploadDict.Add "Pack", 0 ' �������� ������ ��
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

' ������� ��� ����������� ���������� ����� �� �����
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

' ������ ��� ��������� ���������� ������: 
' ������� ������ ����� � ������� �������� ��������
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

' �������� ������ �� INI-�����
' ResDict - ������ Dictionary, ��� �������� ���� ����/��������
Function GetDataFromIniFile(ByVal IniFileName, ByRef ResDict)
      GetDataFromIniFile = false
  
    ' ����� �������������
    Dim IniFile 'As TextStream

    On Error Resume Next
    Dim ForRead
    ForRead =1
    Set IniFile = fso.OpenTextFile(IniFileName,ForRead)
    if err.Number<>0 then
      Err.Clear()
      echo "Ini-���� "& IniFileName &" �� ������� �������!"
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
    ' ���� �� ������-�����������  
      if not RegExpTest("^\s*[;']",s) then
    '  For index=0 To IniDict.Count-1
    '    reg.Pattern="\s*"+elem(index)+"\s*=\s*(.+)"
    ' �������� ���� � �������� � Ini-�����, ����� ��������� �����������
        Set Matches = reg.Execute(s)
        if Matches.Count>0 then
   
        Dim lkey, lvalue
		' �������� ����� ����, �������� �� �������� ��������� � �����(� ������) �������    
					'ResDict.Add elem(index),Trim(replace(Matches(0).SubMatches(0),vbTab," "))
            lkey = LCase(Trim(replace(Matches(0).SubMatches(0),vbTab," ")))
            lvalue = replace(Matches(0).SubMatches(1), vbTab, " ")
            lvalue = Trim(replace(lvalue, chr(34), "")) '������ �������
            
            ResDict.Add lkey, lvalue
					'ResDict.Add LCase(Trim(replace(Matches(0).SubMatches(0),vbTab," "))),Trim(replace(Matches(0).SubMatches(1),vbTab," "))

Debug "lkey=lvalue", lkey + " = [" + lvalue + "]"
        end if
      end if
    Loop
    IniFile.Close()

    if ResDict.Count=0 then
      echo "�� ������� �������� ������ �� Ini-����� " & IniFileName
      GetDataFromIniFile = false
    else  
      GetDataFromIniFile = true
    end if
End Function 'GetDataFromIniFile


' ��������� �� ������������ �������
' ������� �������� �� �����
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
'    Wscript.Echo "�������� ���� �� ���� A:. ��� �� ��������� ����� �����."
    Wscript.Echo "��������� ���������������� �������� � ����� 1� 8.2"
    Wscript.Echo ""
    Wscript.Echo "��������� ������:"
    Wscript.Echo "  "+ WScript.ScriptName +" [����-�������� | /? | /h]"
    Wscript.Echo ""
    Wscript.Echo "������:"
    Wscript.Echo "1. cscript "+ WScript.ScriptName +" ����.ini"
    Wscript.Echo "2. cscript "+ WScript.ScriptName
    Wscript.Echo "   ���������� ���� �����."

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
    If blnErrorOccurred("���������� �������") Then
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
