Option Explicit
Dim objFSO, objFSOText, objFolder, objFile
Dim strDirectory, strFile
strDirectory = "E:\logs\guy1"
strFile = "\Summer.txt"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create the Folder specified by strDirectory on line 10
Set objFolder = objFSO.CreateFolder(strDirectory)

' -- The heart of the create file script
'-----------------------
'Creates the file using the value of strFile on Line 11
' -----------------------------------------------
Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
Wscript.Echo "Just created " & strDirectory & strFile

Wscript.Quit