' Starts the Teams Calendar Helper (published ICS link) with no console window.
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
CreateObject("WScript.Shell").Run "node """ & here & "\ics-calendar-helper.mjs"" serve", 0, False
