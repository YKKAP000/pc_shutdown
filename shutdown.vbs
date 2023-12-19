Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

' カスタム フォームを作成
strMessage = "終業ボタンを押しましたか?" & vbCrLf & vbCrLf & "Yes: シャットダウン開始" & vbCrLf & "No: シャットダウン中止. Teamsを開きます"

' カスタム フォームを表示
result = objWshShell.Popup(strMessage, 0, "確認", vbYesNo + vbQuestion)

If result = vbYes Then
    ' Yesが選択された場合、シャットダウンを実行
    objWshShell.Run "C:\Windows\system32\shutdown.exe -s -f -t 60"
Else
    ' Noが選択された場合、シャットダウンを取りやめてTeamsを起動
    objWshShell.Run "C:\Users\VAIO\AppData\Local\Microsoft\Teams\Update.exe --processStart ""Teams.exe""" ' Teamsの実行可能ファイルのパスを指定
    WScript.Quit
End If
