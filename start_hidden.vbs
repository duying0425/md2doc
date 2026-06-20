Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

root = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = root

Set env = shell.Environment("PROCESS")
srcPath = root & "\src"
If Len(env("PYTHONPATH")) > 0 Then
    env("PYTHONPATH") = srcPath & ";" & env("PYTHONPATH")
Else
    env("PYTHONPATH") = srcPath
End If

pythonPath = root & "\.venv\Scripts\python.exe"
If fso.FileExists(pythonPath) Then
    shell.Run """" & pythonPath & """ -m md2doc", 0, False
Else
    shell.Run "python.exe -m md2doc", 0, False
End If
