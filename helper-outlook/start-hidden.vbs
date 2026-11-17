' Starts the Teams Calendar Helper v2 (classic Outlook) with no console window.
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
CreateObject("WScript.Shell").Run "node """ & here & "\outlook-calendar-helper.mjs"" serve", 0, False
