dim RegEx
Set RegEx = New RegExp
With RegEx
      .Pattern = "^F[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]\.txt$"
      .IgnoreCase = True
      .Global = False
End With

'Add Clients String:
MsgBox  RegEx.test("F12345678")
MsgBox  RegEx.test("F12345678.txt")

If RegEx.test("F12345678.txt") Then
    MsgBox "Hey!"
End if  


'FSO.GetFile("a.txt").Name = "b.txt"

'
'Fso.MoveFile "A.txt", "B.txt"
