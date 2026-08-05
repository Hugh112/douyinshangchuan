Option Explicit
Dim shell, fso, baseDir, exePath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
exePath = fso.BuildPath(baseDir, "DouyinPublisher.exe")
If fso.FileExists(exePath) Then
  shell.Run Chr(34) & exePath & Chr(34), 0, False
End If
