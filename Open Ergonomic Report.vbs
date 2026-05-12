Option Explicit

Dim shell, fso, projectRoot, electronPath

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectRoot = fso.GetParentFolderName(WScript.ScriptFullName)
electronPath = projectRoot & "\node_modules\electron\dist\electron.exe"

shell.CurrentDirectory = projectRoot
On Error Resume Next
shell.Environment("PROCESS").Remove "ELECTRON_RUN_AS_NODE"
On Error GoTo 0

shell.Run """" & electronPath & """ .", 1, False
