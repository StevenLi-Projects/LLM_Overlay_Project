Option Explicit

Dim fso, shell, scriptDir, psScript, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psScript = fso.BuildPath(scriptDir, "NotepadMarkdownAssistant.ps1")

If Not fso.FileExists(psScript) Then
    MsgBox "NotepadMarkdownAssistant.ps1 was not found next to Launch-Assistant.vbs.", vbCritical, "Local Text Formatting Assistant"
    WScript.Quit 1
End If

command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & psScript & Chr(34)
shell.CurrentDirectory = scriptDir

' Window style 0 keeps PowerShell hidden. False returns immediately so the tray app keeps running.
shell.Run command, 0, False
