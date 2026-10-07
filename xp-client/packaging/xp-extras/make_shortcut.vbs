' ---------------------------------------------------------------------------
' ORT XP lite client - create a desktop shortcut (double-click me once).
'
' ASCII-ONLY on purpose: Windows Script Host is present on Windows XP by default
' (PowerShell is NOT), and Chinese text in a .vbs would need the right ANSI code page.
'
' Creates "ORT-XP.lnk" on the current user's desktop, pointing at ORT-XP.exe in this
' folder and setting the working directory to this folder (the client resolves its
' own Data\local_settings.json relative to the program directory).
' ---------------------------------------------------------------------------
Option Explicit

Dim fso, shell, scriptDir, target, linkPath, link

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
target = fso.BuildPath(scriptDir, "ORT-XP.exe")

If Not fso.FileExists(target) Then
    MsgBox "ORT-XP.exe was not found next to this script." & vbCrLf & vbCrLf & _
           "Put this script into the ORT-XP folder and run it again.", _
           16, "ORT-XP"
    WScript.Quit 1
End If

linkPath = fso.BuildPath(shell.SpecialFolders("Desktop"), "ORT-XP.lnk")
Set link = shell.CreateShortcut(linkPath)
link.TargetPath = target
link.WorkingDirectory = scriptDir
link.Description = "ORT lab management system - XP lite client"
link.Save

MsgBox "Desktop shortcut created:" & vbCrLf & linkPath & vbCrLf & vbCrLf & _
       "Target: " & target, 64, "ORT-XP"
