Set WshShell = createobject("wscript.Shell")

x=Msgbox("To stop spam run wscriptkiller.exe or wscriptkiller.bat",0,"END") 


WScript.Sleep 2000
Do
WScript.Sleep 1000
WshShell.SendKeys "SPAM"
WshShell.SendKeys "~"
Loop