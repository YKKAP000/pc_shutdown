Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

' �J�X�^�� �t�H�[�����쐬
strMessage = "�I�ƃ{�^���������܂�����?" & vbCrLf & vbCrLf & "Yes: �V���b�g�_�E���J�n" & vbCrLf & "No: �V���b�g�_�E�����~. Teams���J���܂�"

' �J�X�^�� �t�H�[����\��
result = objWshShell.Popup(strMessage, 0, "�m�F", vbYesNo + vbQuestion)

If result = vbYes Then
    ' Yes���I�����ꂽ�ꍇ�A�V���b�g�_�E�������s
    objWshShell.Run "C:\Windows\system32\shutdown.exe -s -f -t 60"
Else
    ' No���I�����ꂽ�ꍇ�A�V���b�g�_�E��������߂�Teams���N��
    objWshShell.Run "C:\Users\VAIO\AppData\Local\Microsoft\Teams\Update.exe --processStart ""Teams.exe""" ' Teams�̎��s�\�t�@�C���̃p�X���w��
    WScript.Quit
End If
