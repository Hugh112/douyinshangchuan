Option Explicit
Dim shell, fileSystem, projectRoot
Set shell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")
projectRoot = fileSystem.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = projectRoot
shell.Run "pyw -3 """ & projectRoot & "\app\bitbrowser_setup.pyw""", 1, False
