Option Compare Database
Option Explicit

Public Sub ReadFileFolder(stSpectrumTypeLetter As String)
Dim rstParameter As Recordset, stSpectrumDir As String, intHowLongDays As Integer, stPathByType As String
Dim stMailSubject As String, stMailMessage As String
Dim strIDNUMBER As String, strComment As String
Dim rstUpdateRow As Recordset
Dim strStep As String


On Error GoTo ErrHandler

strStep = "1.main_param"

'������ ����������
Set rstParameter = CurrentDb.OpenRecordset("select Name, Value from Parameter", dbOpenSnapshot)
rstParameter.OpenRecordset
While Not rstParameter.EOF

    If rstParameter!Name = "SpectrumFolder" Then stSpectrumDir = rstParameter!Value
    If rstParameter!Name = "HowLongDays" Then intHowLongDays = rstParameter!Value

    rstParameter.MoveNext
Wend
rstParameter.Close

strStep = "2.block_param"

'�� ���� ������� ���������� �������������� ���� � ������
Set rstParameter = CurrentDb.OpenRecordset("select DirectoryPath from PathByType where Type = """ & stSpectrumTypeLetter & """", dbOpenSnapshot)
rstParameter.OpenRecordset
If Not rstParameter.EOF Then
    stPathByType = rstParameter!DirectoryPath
    Else
    rstParameter.Close
    GoTo NoTypeParam
    End If
rstParameter.Close

strStep = "3.filesystem"

' ������ � �������� ������� � ������� ���������� ������
' � ���������� ������ ������ ��������
CurrentDb.Execute "delete from SpFileList"
   Dim fs As Object _
        , f As Object _
        , f1 As Object _
        , fc _
        , fls _
        , fi _
        , s
    Set fs = CreateObject("Scripting.FileSystemObject")
    strStep = "4.GetFolder: " & stSpectrumDir & stPathByType
    Set f = fs.GetFolder(stSpectrumDir & stPathByType)
    Set fc = f.SubFolders
    For Each f1 In fc
        If f1.DatelastModified > Now() - intHowLongDays Then
            strStep = "5.GetFolder"
            Set f = fs.GetFolder(stSpectrumDir & stPathByType & "\" & f1.Name)
            Set fls = f.Files
            For Each fi In fls
                strStep = "6.GetFile"
                If UCase(Mid(fi.Name, 1, 1)) = "F" Then
                    '��� �� ����� ��� ����� � ��� ����� �������
                    strIDNUMBER = UCase(Left(fi.Name, 10))
                    strComment = Mid(fi.Name, 12, InStr(fi.Name, ".") - 12)
                    CurrentDb.Execute "insert into SpFileList(SpectrumType,BlockCode,FileName,IDNUMBER,Comment) " & _
                                      "values(""" & stSpectrumTypeLetter & _
                                                """, """ & f1.Name & _
                                                """, """ & fi.Name & _
                                                """, """ & strIDNUMBER & _
                                                """, """ & strComment & """ )"
                    
                    's = s & f1.Name & "\" & fi.Name
                    's = s & vbCrLf
                End If
            Next
        End If
    Next

strStep = "7.Update"

 'intHowLongDays = intHowLongDays / (intHowLongDays - intHowLongDays) ' ����������� ������
CurrentDb.Execute "UPDATE SpFileList SET SpFileList!Percent = Comment WHERE IsNumeric(Comment)=True;"
CurrentDb.Execute "UPDATE SpFileList SET SpFileList!Answer = UCase(Comment) WHERE IsNull(Percent)=True;"

strStep = "8.CheckResult"

' ���� ���� ���������� ��������� - ����� ����� ���������
Set rstUpdateRow = CurrentDb.QueryDefs("PreUpdateResult").OpenRecordset
If Not rstUpdateRow.EOF Then
    While Not rstUpdateRow.EOF
        '������� � �������
       Call WriteLog("���� " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; ��������� = " & rstUpdateRow!Answer)
       rstUpdateRow.MoveNext
    Wend
        '� ������
    CurrentDb.QueryDefs("UpdateResult").Execute
End If
rstUpdateRow.Close

strStep = "9.CheckPercent"
' ���� ������
Set rstUpdateRow = CurrentDb.QueryDefs("PreUpdatePercent").OpenRecordset
If Not rstUpdateRow.EOF Then
    While Not rstUpdateRow.EOF
        '������� � �������???
       Call WriteLog("���� " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; ���������� �� ������� = " & rstUpdateRow![����������_��_�������] & "; ������� = " & rstUpdateRow!Percent & "; ��������� = OK")
       rstUpdateRow.MoveNext
    Wend
        '� ������
    CurrentDb.QueryDefs("UpdatePercentResult").Execute
End If
rstUpdateRow.Close

strStep = "10.ChangeResult"

' ���� ���������� ��������� �� ������ � ����������� �������
Set rstUpdateRow = CurrentDb.QueryDefs("PreChangeResult").OpenRecordset
If Not rstUpdateRow.EOF Then
    stMailMessage = ""
    While Not rstUpdateRow.EOF

        '������� � �������
       Call WriteLog("���� " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; ��������� ��������� � " & rstUpdateRow![���������] & " �� " & rstUpdateRow!Answer)
        '�������� �� �����
       stMailMessage = stMailMessage & "���� " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; ��������� ��������� � " & rstUpdateRow![���������] & " �� " & rstUpdateRow!Answer & vbCrLf

       rstUpdateRow.MoveNext
    Wend
    Call SendMailMessage("������ ��������� �������", stMailMessage)
        '� ������
    CurrentDb.QueryDefs("UpdateChangeResult").Execute
End If
rstUpdateRow.Close

' ���� ���������� ��������� �� ������ � ����������� �������
Set rstUpdateRow = CurrentDb.QueryDefs("PreChangePercent").OpenRecordset
If Not rstUpdateRow.EOF Then
    stMailMessage = ""
    While Not rstUpdateRow.EOF

        '������� � �������
       Call WriteLog("���� " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; ������� ���������� � " & rstUpdateRow![���������� ����������] & " �� " & rstUpdateRow!Percent)
        '�������� �� �����
       stMailMessage = stMailMessage & "���� " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; ������� ���������� � " & rstUpdateRow![���������� ����������] & " �� " & rstUpdateRow!Percent & vbCrLf

       rstUpdateRow.MoveNext
    Wend
    Call SendMailMessage("������ ������� ������� �������", stMailMessage & "������� �� �������� �� ��������!")
        '������ �� ������!
End If
rstUpdateRow.Close




Exit Sub

NoTypeParam:
'��� ���-�� ����� �������� � ������ ��� ���������� ��������� ��� ����
Call WriteLog("��� ���� �������� " & stSpectrumTypeLetter & " �� ������ ���� � ��������� ������.")
Exit Sub

ErrHandler:

Call WriteLog("��� �������� ������ ���� " & stSpectrumTypeLetter & " ��������� ������. ��� " & strStep & " ��� = " & Err.Number & ": " & Err.Description)
' ������������ " & Environ("USERNAME") & "
End Sub
'����� ����� ���� ������� � �������
Public Function CheckSpectrumFolders(stSpectrumTypeLetter As String) As Integer

Call ReadFileFolder(stSpectrumTypeLetter)

CheckSpectrumFolders = 0

End Function
'����� ������ �������� � �������, ��������� � ����������. ���� ��� ����� ��� �������� ��� - ����� � ������ ��
Function WriteLog(Optional stMessage As String = "") As Integer
Dim rstParameter As Recordset, stLogPath As String, fs As Object

Set fs = CreateObject("Scripting.FileSystemObject")

Set rstParameter = CurrentDb.OpenRecordset("SELECT IIf(IsNull([Value]),"""",[Value]) AS NValue FROM Parameter WHERE (Parameter.Name=""LogPath"")", dbOpenSnapshot)
rstParameter.OpenRecordset
If Not rstParameter.EOF Then
    stLogPath = rstParameter!NValue
    stLogPath = IIf(fs.FolderExists(stLogPath), stLogPath, "")
    Else: stLogPath = ""
    End If
rstParameter.Close

stLogPath = IIf(stLogPath = "", CurrentProject.Path, stLogPath)

Open stLogPath & "\" & CurrentProject.Name & "_log.txt" For Append As #1
Print #1, Now(); stMessage
Close #1

WriteLog = 0

End Function
'�������� ��������� �� ��.�����
Function SendMailMessage(stSubject As String, stBody As String) As Boolean
SendMailMessage = False
On Error GoTo ErrExit

Dim rstParameter As Recordset
Dim stMailRecipients As String

Set rstParameter = CurrentDb.OpenRecordset("select Name, Value from Parameter where Name = ""MailRecipients""", dbOpenSnapshot)
rstParameter.OpenRecordset
If Not rstParameter.EOF Then stMailRecipients = rstParameter!Value
rstParameter.Close

Call ExecSQL("exec sendMessage '" & stMailRecipients & "', '" & stSubject & "', '" & stBody & "'")

SendMailMessage = True
Exit Function

ErrExit:
SendMailMessage = False
End Function

' ��� ������ SQL ��������
Public Function ExecSQL(sql As String) As Boolean
On Error GoTo ErrExit

Static qd As QueryDef
    
    If qd Is Nothing Then
        Set qd = CurrentDb.QueryDefs("ExecSQL")
    End If

    qd.sql = sql
    qd.ODBCTimeout = 1000
    qd.ReturnsRecords = False
    qd.Connect = CurrentDb().TableDefs("dbo_tbl������").Connect
    qd.Execute

    ExecSQL = True

    Exit Function

ErrExit:
 ExecSQL = False
 Call WriteLog("������ ��� ������ ExecSQL. ��� = " & Err.Number & ": " & Err.Description)
End Function