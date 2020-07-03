'������:
'cscript.exe chest_start.vbs au_BDB.ini %USERPROFILE%

Option Explicit 
Const ForReading = 1 
Const ForWriting = 2
Const ChestPath = "\!Chest35"


Sub Print(x)
	WScript.Echo x
End Sub

Sub Main

  Dim IniFileName, Master, Slave, DateTime, AutoRun, FSO, IniFile, TargetFolderName, TargetFolder, EqlSignPos, MasterFile, ReadIniFile, TmpIniFile

  if (wscript.arguments.count > 2 or wscript.arguments.count < 1) then
    Print("Usage: chest_start.vbs <program.ini> <root_folder>")
    wscript.quit 1
  else
    IniFileName = wscript.arguments.Item(0)  
  end if

  TargetFolderName = ""
  if wscript.arguments.count = 2 then TargetFolderName =  wscript.arguments.Item(1)&ChestPath                ' ���� �� ������ ��������� ������ �������, ������� ����� ��� ���������� !Chest35

  Set FSO = CreateObject("Scripting.FileSystemObject")

  if  TargetFolderName <> ""   then                                                                           
	if not FSO.FolderExists(TargetFolderName) then FSO.CreateFolder TargetFolderName
	Set TargetFolder =  FSO.GetFolder(TargetFolderName) ' "%USERPROFILE%"
	if  not FSO.FileExists(TargetFolderName&"\"&IniFileName) then
		if not FSO.FileExists(IniFileName) then 
			wscript.quit 2	
		else
 			Set IniFile = FSO.GetFile(IniFileName)
			FSO.CopyFile IniFile.Path, TargetFolderName&"\"
 			Set IniFile = FSO.GetFile(TargetFolderName&"\"&IniFileName)
		end if		
  		
	end if
	Set IniFile = FSO.GetFile(TargetFolderName&"\"&IniFileName)
  else
	if not FSO.FileExists(IniFileName) then wscript.quit 2  else Set IniFile = FSO.GetFile(IniFileName)    ' ���������� ini-���� �� �������� ��������
	TargetFolderName = FSO.GetParentFolderName(IniFile)                                                    ' ������� ���������� ������� �������(???)
	Set TargetFolder =  FSO.GetFolder(TargetFolderName)
  end if

  Set ReadIniFile = FSO.OpenTextFile(IniFile.Path, ForReading)							' ��������� ini-���� ��� ���������� ����������

  Dim str
  Do while not ReadIniFile.AtEndofStream                								'  ��������� ������� � ����������
    str=ReadIniFile.ReadLine
    'Print(str)
    EqlSignPos = InStr(str,"=")
    if EqlSignPos > 0 then '
	if left(str,EqlSignPos-1) = "Master" then Master = mid(str,EqlSignPos+1)				' ���� � ���������� �����
	if left(str,EqlSignPos-1) = "Slave" then Slave = mid(str,EqlSignPos+1)                                  ' �� ������������
	if left(str,EqlSignPos-1) = "DateTime" then DateTime = mid(str,EqlSignPos+1)  				' ������� ���� ����������� ���������
	if left(str,EqlSignPos-1) = "AutoRun" then AutoRun = mid(str,EqlSignPos+1)                              ' ����� �� ��������� ����� ����������
	end if	
  Loop
  ReadIniFile.Close

if FSO.FileExists(Master) then Set MasterFile = FSO.GetFile(Master)  else wscript.quit 3

Set IniFile = FSO.GetFile(TargetFolderName&"\"&IniFileName)
if CDate(DateTime) < MasterFile.DateLastModified or not FSO.FileExists(TargetFolder.Path&"\"&MasterFile.Name) then 
	Print("Need copy!") 
	MasterFile.Copy  TargetFolder.Path&"\"&MasterFile.Name

	set TmpIniFile = FSO.CreateTextFile (left(IniFile.Path,len(IniFile.Path)-1)&"_", ForWriting)             ' ���������� �� �� �����, ����� ����
        TmpIniFile.WriteLine("[Update]")
        TmpIniFile.WriteLine("Master="&Master)
        TmpIniFile.WriteLine("Slave="&Slave)
        TmpIniFile.WriteLine("DateTime="&MasterFile.DateLastModified)
        TmpIniFile.WriteLine("AutoRun="&AutoRun)
	TmpIniFile.Close

	Set TmpIniFile = FSO.GetFile(left(IniFile.Path,len(IniFile.Path)-1)&"_")
	TmpIniFile.Copy IniFile.Path
	TmpIniFile.Delete	
else 
	Print("Ready!")
END IF

Dim WSShell
if AutoRun = "True" then 
	set WSShell = CreateObject("WScript.Shell")
	WSShell.Run TargetFolder.Path&"\"&MasterFile.Name, 2 , false
end if

End Sub

Main