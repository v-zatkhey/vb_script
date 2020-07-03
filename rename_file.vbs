Dim Fso, oFolder, oFile, RegEx 

Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
set oFolder = Fso.GetFolder(".")
 
Set RegEx = New RegExp
With RegEx
      .Pattern = "^F[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]\.mol$"
      .IgnoreCase = True
      .Global = False
End With


for each oFile in oFolder.Files

	If RegEx.test(oFile.Name) Then
	    oFile.Name = left(oFile.Name,5) & "-" & mid(oFile.Name,6)
	End if  
	
next

