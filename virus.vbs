Option Explicit
Dim objFSO, objFolder, objShell, objFile
Dim strDirectory, strFile
strDirectory = "e:\logs"
strFile = "\Virus.txt"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check that the strDirectory folder exists
If objFSO.FolderExists(strDirectory) Then
Set objFolder = objFSO.GetFolder(strDirectory)
Else
Set objFolder = objFSO.CreateFolder(strDirectory)
WScript.Echo "Just created " & strDirectory
End If

If objFSO.FileExists(strDirectory & strFile) Then
Set objFolder = objFSO.GetFolder(strDirectory)
Else
Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
Wscript.Echo "Just created " & strDirectory & strFile
End If

set objFolder = nothing
set objFile = nothing

If err.number = vbEmpty then
Set objShell = CreateObject("WScript.Shell")
objShell.run ("Explorer" & " " & strDirectory & "\" )
Else WScript.echo "VBScript Error: " & err.number
End If

WScript.Quit