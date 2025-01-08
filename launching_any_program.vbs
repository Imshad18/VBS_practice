set shell = WScript.CreateObject("WScript.shell")
shell.Run("C:\Windows\System32\calc.exe " & WScript.ScriptFullName),0,True 
'you can replace any program path here
