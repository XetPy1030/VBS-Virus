Dim a

Do
CreateObject("WScript.Shell").Run "notepad.exe"
a = a+1
Loop Until a>200