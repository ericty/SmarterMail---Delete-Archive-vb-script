''
 ' ----------------------------------------------------------------------------
 ' "THE BEER-WARE LICENSE" (Revision 42):
 ' <eric@truenet.com> wrote this file.  As long as you retain this notice you
 ' can do whatever you want with this stuff. If we meet some day, and you think
 ' this stuff is worth it, you can buy me a beer in return.   Eric Tykwinski
 ' ----------------------------------------------------------------------------
''
' Simple script to delete SmarterTools Indexed files for archives.
' Written so we have a way to store archives for X days for free,
' and charge for Yearly or 7 year rentention.
'

Dim fso, filepath, filename, filearray
Set fso = CreateObject("Scripting.FileSystemObject")
'wscript.echo WScript.Arguments.Count
If WScript.Arguments.Count <> 2 Then
	wscript.echo "Wrong format: DeleteArchiveAfterXdays.vbs (FolderPath in Quotes) Number"
Else
	Dim RootFolder
	Dim Days
	RootFolder = WScript.Arguments.Item(0)
	Days = CInt(WScript.Arguments.Item(1))
	
	'Get SubFolders
	Call ShowSubFolders(fso.GetFolder(RootFolder), Days)
End If
Set fso = Nothing

Sub ShowSubFolders(RootFolder, Days)
	Dim SubFolder
	Set re = New RegExp
	re.Pattern = "\\Indexed$"
	For Each Subfolder in RootFolder.SubFolders
		If re.Test(Subfolder.Path) <> false Then
			'Wscript.Echo Subfolder.Path
			Call DeleteFiles(Subfolder.Path, Days)
		End If
		ShowSubFolders Subfolder, Days
	Next
End Sub

Sub DeleteFiles (FolderName, Days)
	Set filepath = fso.GetFolder(FolderName)
	Set filearray = filepath.Files
	For Each filename in filearray
	'wscript.echo filename,"modified",filename.DateLastModified
	If DateDiff("d", filename.DateLastModified, Now) > Days Then
		fso.DeleteFile filename
		'wscript.echo filename,"deleted"
	End If
	'wscript.echo DateDiff("d", filename.DateLastModified, Now)
	Next

	Set filepath = Nothing
	Set filearray = Nothing 
End Sub