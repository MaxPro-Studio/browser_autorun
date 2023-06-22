Set oShell = CreateObject("WScript.Shell")
oShell.Run "chrome.exe"
WScript.Sleep 500
oShell.SendKeys "% n"
oShell.SendKeys("% {DOWN}{DOWN}{DOWN}{DOWN}{ENTER}")