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

'чтение параметров
Set rstParameter = CurrentDb.OpenRecordset("select Name, Value from Parameter", dbOpenSnapshot)
rstParameter.OpenRecordset
While Not rstParameter.EOF

    If rstParameter!Name = "SpectrumFolder" Then stSpectrumDir = rstParameter!Value
    If rstParameter!Name = "HowLongDays" Then intHowLongDays = rstParameter!Value

    rstParameter.MoveNext
Wend
rstParameter.Close

strStep = "2.block_param"

'по типу спектра определяем дополнительный путь к блокам
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

' роемся в файловой системе в поисках подходящих файлов
' и составляем список файлов спектров
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
                    'тут мы знаем код блока и имя файла спектра
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

 'intHowLongDays = intHowLongDays / (intHowLongDays - intHowLongDays) ' специальная ошибка
CurrentDb.Execute "UPDATE SpFileList SET SpFileList!Percent = Comment WHERE IsNumeric(Comment)=True;"
CurrentDb.Execute "UPDATE SpFileList SET SpFileList!Answer = UCase(Comment) WHERE IsNull(Percent)=True;"

strStep = "8.CheckResult"

' если есть негативный результат - сразу пишем результат
Set rstUpdateRow = CurrentDb.QueryDefs("PreUpdateResult").OpenRecordset
If Not rstUpdateRow.EOF Then
    While Not rstUpdateRow.EOF
        'отметим в журнале
       Call WriteLog("Блок " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; результат = " & rstUpdateRow!Answer)
       rstUpdateRow.MoveNext
    Wend
        'в спектр
    CurrentDb.QueryDefs("UpdateResult").Execute
End If
rstUpdateRow.Close

strStep = "9.CheckPercent"
' если цифрой
Set rstUpdateRow = CurrentDb.QueryDefs("PreUpdatePercent").OpenRecordset
If Not rstUpdateRow.EOF Then
    While Not rstUpdateRow.EOF
        'отметим в журнале???
       Call WriteLog("Блок " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; требования по чистоте = " & rstUpdateRow![Требования_по_чистоте] & "; чистота = " & rstUpdateRow!Percent & "; результат = OK")
       rstUpdateRow.MoveNext
    Wend
        'в спектр
    CurrentDb.QueryDefs("UpdatePercentResult").Execute
End If
rstUpdateRow.Close

strStep = "10.ChangeResult"

' если негативный результат не совпал с результатом спектра
Set rstUpdateRow = CurrentDb.QueryDefs("PreChangeResult").OpenRecordset
If Not rstUpdateRow.EOF Then
    stMailMessage = ""
    While Not rstUpdateRow.EOF

        'отметим в журнале
       Call WriteLog("Блок " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; результат изменился с " & rstUpdateRow![Результат] & " на " & rstUpdateRow!Answer)
        'отправим по почте
       stMailMessage = stMailMessage & "Блок " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; результат изменился с " & rstUpdateRow![Результат] & " на " & rstUpdateRow!Answer & vbCrLf

       rstUpdateRow.MoveNext
    Wend
    Call SendMailMessage("изменён результат спектра", stMailMessage)
        'в спектр
    CurrentDb.QueryDefs("UpdateChangeResult").Execute
End If
rstUpdateRow.Close

' если негативный результат не совпал с результатом спектра
Set rstUpdateRow = CurrentDb.QueryDefs("PreChangePercent").OpenRecordset
If Not rstUpdateRow.EOF Then
    stMailMessage = ""
    While Not rstUpdateRow.EOF

        'отметим в журнале
       Call WriteLog("Блок " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; чистота изменилась с " & rstUpdateRow![Процентное содержание] & " на " & rstUpdateRow!Percent)
        'отправим по почте
       stMailMessage = stMailMessage & "Блок " & rstUpdateRow!BlockCode & " IDNUMBER " & rstUpdateRow!IDNUMBER & "; чистота изменилась с " & rstUpdateRow![Процентное содержание] & " на " & rstUpdateRow!Percent & vbCrLf

       rstUpdateRow.MoveNext
    Wend
    Call SendMailMessage("изменён процент чистоты спектра", stMailMessage & "Решение по спектрам не менялось!")
        'спектр не меняем!
End If
rstUpdateRow.Close




Exit Sub

NoTypeParam:
'тут что-то нужно записать в журнал про отсутствуе параметра для типа
Call WriteLog("Для типа спектров " & stSpectrumTypeLetter & " не указан путь к каталогам блоков.")
Exit Sub

ErrHandler:

Call WriteLog("При проверке блоков типа " & stSpectrumTypeLetter & " произошла ошибка. Шаг " & strStep & " Код = " & Err.Number & ": " & Err.Description)
' Пользователь " & Environ("USERNAME") & "
End Sub
'чтобы можно было вызвать в макросе
Public Function CheckSpectrumFolders(stSpectrumTypeLetter As String) As Integer

Call ReadFileFolder(stSpectrumTypeLetter)

CheckSpectrumFolders = 0

End Function
'пишем журнал действий в каталог, указанный в параметрах. если там пусто или каталога нет - рядом с файлом БД
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
'отправка сообщения по эл.почте
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

' для вызова SQL процедур
Public Function ExecSQL(sql As String) As Boolean
On Error GoTo ErrExit

Static qd As QueryDef
    
    If qd Is Nothing Then
        Set qd = CurrentDb.QueryDefs("ExecSQL")
    End If

    qd.sql = sql
    qd.ODBCTimeout = 1000
    qd.ReturnsRecords = False
    qd.Connect = CurrentDb().TableDefs("dbo_tblСпектр").Connect
    qd.Execute

    ExecSQL = True

    Exit Function

ErrExit:
 ExecSQL = False
 Call WriteLog("Ошибка при вызове ExecSQL. Код = " & Err.Number & ": " & Err.Description)
End Function