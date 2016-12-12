set WshShell = CreateObject("WScript.Shell")
For iCounter = i to 20000
WshShell.SendKeys "{NUMLOCK}"
WScript.Sleep 100
WshShell.SendKeys "{CAPSLOCK}"
WScript.Sleep 100
WshShell.SendKeys "{SCROLLLOCK}"
WScript.Sleep 100
WshShell.SendKeys "{NUMLOCK}"
WScript.Sleep 100
WshShell.SendKeys "{CAPSLOCK}"
WScript.Sleep 100
WshShell.SendKeys "{SCROLLLOCK}"
WScript.Sleep 100
Next