Option Explicit 
Const WorkPath = "D:\work\VBScript"
Const addPath = "\1"


Sub Print(x)
	WScript.Echo x
End Sub

Sub Main

   Dim FSO, WrkFolder, SubFolders, SubFolder, InnerSubFolders, InnerSubFolder

   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   Set WrkFolder = FSO.GetFolder(WorkPath)
   Set SubFolders = WrkFolder.SubFolders

 
   If SubFolders.Count <> 0 Then
      For Each SubFolder In SubFolders

		
	Print SubFolder.Path
	if not (FSO.FolderExists(SubFolder.Path & addPath)) then FSO.CreateFolder SubFolder.Path & addPath 		

	Set InnerSubFolders = SubFolder.SubFolders
	for each InnerSubFolder in InnerSubFolders
		
	        if InnerSubFolder.Name <> "1" then
			Print InnerSubFolder.Path & " --> " & SubFolder.Path & addPath & "\" & InnerSubFolder.Name 
			FSO.MoveFolder  InnerSubFolder.Path, SubFolder.Path & addPath & "\" & InnerSubFolder.Name 
		end if

	Next

	Print SubFolder.Path & "\*.*" & " --> " & SubFolder.Path & addPath  
	FSO.MoveFile  SubFolder.Path & "\*.*", SubFolder.Path & addPath

      Next
   End If

End Sub

Main