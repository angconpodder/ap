'Script to get MD5 Signature and Size in Bytes
'Created by: Angcon Podder
'Last Updated: 3-22-2018
'Last Updated by: Angcon

set objShell = WScript.CreateObject ("WScript.Shell")
set objFSO = createobject("Scripting.FileSystemObject")

fcivExe = objShell.CurrentDirectory & "\FCIV\fciv.exe"
outFile = objShell.CurrentDirectory & "\fileInfo.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)

strFile = objShell.CurrentDirectory + "\DropFileHere\"
Set objFolder = objFSO.GetFolder(strFile)
Set colFiles = objFolder.Files

Count = 0
For Each obFile in colFiles
	
	strPath = strFile + obFile.Name
	objExecCmd = "cmd.exe /c " & fcivExe & " " & chr(34) & strPath & chr(34)
	objFile.Write obFile.Name & vbCrLf
	objFile.Write "File Size (bytes): " & objFSO.GetFile(strPath).Size & vbCrLf
	
	Set ObjExec = objShell.Exec(objExecCmd)
	
	Do
		count = count + 1
		strFromProc = ObjExec.StdOut.ReadLine()
		if count = 4 Then
			sig = split(strFromProc)
			objFile.Write "File Signature: " & sig(0) & vbCrLf
		end if
	Loop While Not ObjExec.Stdout.atEndOfStream
		
	objFile.Write "===============" & vbCrLf
	count = 0
Next


objFile.Close

'objShell.run "C:\Users\APODDER\Desktop\FileInfo\FCIV\fciv.exe " & outFile
objShell.run "notepad.exe " & outFile

Set objFile = nothing
Set colFiles = nothing
Set objFolder = nothing
set objFSO = nothing
set objShell = nothing

