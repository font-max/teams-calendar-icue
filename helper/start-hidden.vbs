' Starts the Teams Calendar Helper with no console window.
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
CreateObject("WScript.Shell").Run "node """ & here & "\teams-calendar-helper.mjs"" serve", 0, False
