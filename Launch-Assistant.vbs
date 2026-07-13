Option Explicit

Dim fso, shell, scriptDir, appPath, psScript, configPath, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appPath = fso.BuildPath(scriptDir, "dist\LocalTextFormattingAssistant.exe")
psScript = fso.BuildPath(scriptDir, "NotepadMarkdownAssistant.ps1")
configPath = fso.BuildPath(scriptDir, "config.json")

If fso.FileExists(appPath) Then
    command = Chr(34) & appPath & Chr(34) & " --config " & Chr(34) & configPath & Chr(34)
Else
    If Not fso.FileExists(psScript) Then
        MsgBox "Neither the compiled app nor NotepadMarkdownAssistant.ps1 was found.", vbCritical, "Local Text Formatting Assistant"
        WScript.Quit 1
    End If
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & psScript & Chr(34)
End If
shell.CurrentDirectory = scriptDir

' Window style 0 keeps the legacy fallback hidden. False returns immediately.
shell.Run command, 0, False
