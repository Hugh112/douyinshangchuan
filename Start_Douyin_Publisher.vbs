Option Explicit
Dim fso, shell, here, appDir, appPath, rc, lastErr
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
If fso.FileExists(fso.BuildPath(here, "app_main.pyw")) Then
    appDir = here
ElseIf fso.FolderExists(fso.BuildPath(here, "app")) Then
    appDir = fso.BuildPath(here, "app")
Else
    MsgBox "Cannot find app folder:" & vbCrLf & here, 48, "Start failed"
    WScript.Quit 1
End If

appPath = fso.BuildPath(appDir, "app_main.pyw")
If Not fso.FileExists(appPath) Then
    MsgBox "Cannot find application file:" & vbCrLf & appPath, 48, "Start failed"
    WScript.Quit 1
End If

shell.CurrentDirectory = appDir
shell.Environment("PROCESS")("PYTHONUTF8") = "1"
shell.Environment("PROCESS")("PYTHONIOENCODING") = "utf-8"

' Prefer the windowless Python launcher. Fall back to hidden py/python launchers.
On Error Resume Next
Err.Clear
rc = shell.Run("pyw.exe " & QuoteArg(appPath), 0, False)
lastErr = Err.Number
If lastErr <> 0 Then
    Err.Clear
    rc = shell.Run("pythonw.exe " & QuoteArg(appPath), 0, False)
    lastErr = Err.Number
End If
If lastErr <> 0 Then
    Err.Clear
    rc = shell.Run("py.exe " & QuoteArg(appPath), 0, False)
    lastErr = Err.Number
End If
If lastErr <> 0 Then
    Err.Clear
    rc = shell.Run("python.exe " & QuoteArg(appPath), 0, False)
    lastErr = Err.Number
End If
On Error GoTo 0

If lastErr <> 0 Then
    MsgBox "Python could not be started. Please install Python and required dependencies." & vbCrLf & _
           "Application: " & appPath, 48, "Start failed"
    WScript.Quit 1
End If

Function QuoteArg(value)
    QuoteArg = Chr(34) & value & Chr(34)
End Function
