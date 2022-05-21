Option Explicit
Dim folder, subFolders, objFF, objFso, objFolder, sameDrive, fromDriveNotFound, toDriveNotFound, startFolder, list, usrName, destinationFolder, objShell, fromDrive, toDrive

Set list = CreateObject("System.Collections.ArrayList")
'This foldernames will be ignored
list.Add "AppData"
'list.Add "Dokumente" 'DE
'list.Add "Documents" 'EN

fromDriveNotFound = True
toDriveNotFound = True
sameDrive = True

'change to foldername if you dont want to start in root directory
startFolder = ""

'select folder, return path
Function SelectFolder(strText)
	Set objShell  = CreateObject( "Shell.Application" )
	Set objFolder = objShell.BrowseForFolder( 0, strText, 0, startFolder )
	If (Not objFolder Is Nothing) then
		SelectFolder = objFolder.Self.Path
	Else
		Set objFolder = nothing
	End If
End Function

'copy files except listed
Function CopyFiles(fromDrive, toDrive)
	Dim file, element
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set fromDrive = objFso.GetFolder(fromDrive)
	Set subFolders = fromDrive.SubFolders
	Set toDrive = objFso.GetFolder(toDrive)
	
	'Copy files
	For Each file in fromDrive.Files
		objFso.CopyFile file, (toDrive.Path & "\")
	Next
	'copy Folders and subElements
	For Each folder in subFolders
	
		Do
			For Each element in list
				If folder.Name = element Then Exit Do
			Next
			objFso.CopyFolder folder , (toDrive.Path & "\")
		Loop While False
	Next
	'uncheck to copy all
	'objFso.CopyFolder fromDrive, toDrive 
End Function

'copy bookmarks to Bookmarks folder
Function CopyBookmarks(usrName)
	Dim firefox, subfirefox, subStringFF
	If usrName = "" Then
		Exit Function
	End If
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFF = objFso.GetFolder(destinationFolder)
	MsgBox(objFF)
	'Create new directory
	If Not objFso.FolderExists(destinationFolder & "\Bookmarks") Then
		objFso.CreateFolder((destinationFolder & "\Bookmarks\"))
	End If
	''Edge Bookmarks
	objFso.CopyFile ("C:\Users\" & usrName & "\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks"), (destinationFolder & "\Bookmarks\")

End Function


Do While sameDrive
	
	Do While fromDriveNotFound
		fromDrive = SelectFolder("Enter a folder/drive from where you want to copy." & vbCrLf & vbCrLf & "[INPUT] --> destination" & vbCrLf)
		If (fromDrive = "") Then 'Exit Code
			WScript.Quit
		Else
			fromDriveNotFound = False
		End If
	Loop

	Do While toDriveNotFound
		toDrive = SelectFolder("Enter a destination folder." & vbCrLf & vbCrLf & fromDrive & " --> [INPUT]" & vbCrLf)
		If ( toDrive = "") Then 'Exit Code
			WScript.Quit
		Else 
			toDriveNotFound = False
			destinationFolder = toDrive
		End If
	Loop
	If fromDrive <> toDrive Then
		sameDrive = False
		MsgBox("Start Copy from " & vbCrLf & vbCrLf & vbTab & fromDrive & _ 
			vbCrLf & vbCrLf & "to" & vbCrLf & vbCrLf & vbTab & toDrive)
		'copy files
		
		'ask for Bookmarks Copy
		gNumber = InputBox("Enter the ProfileName if you want to copy Bookmarks or press Cancel", "Bookmarks")
		if Not gNumber = "" Then 
			CopyBookmarks(gNumber)
		End If
		'start copy files
		CopyFiles fromDrive, toDrive
	Else
		fromDriveNotFound = True
		toDriveNotFound = True
		If Not (IsEmpty(fromDrive) And  IsEmpty(toDrive)) Then
			MsgBox("Sorry, it is impossible to copy files from and to the same folder.") 'tester
		End If
	End If
	
Loop
