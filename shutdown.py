import os
import subprocess
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtGui import QFont

class CustomDialog(QMessageBox):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("お疲れ様でした")
        self.setText("終業ボタンは押し忘れてませんか?")
        self.setInformativeText("No を押すとシャットダウンを\n中止してTeamsを開きます")
        self.setIcon(QMessageBox.Question)

        # スタイルシートを適用
        self.setStyleSheet("""
            QMessageBox {
                background-color: #F0F0F0;  /* グレー基調の背景色 */
                color: #333333;  /* テキストカラー */
                border: 2px solid #B0B0B0;  /* ボーダーカラー */
                border-radius: 10px;
            }
            QLabel {
                color: #333333;  /* テキストカラー */
                padding: 10px;
                margin: 0;
                font-size: 12pt;
                text-align: center;  /* メッセージを中央寄せ */
            }
            QPushButton {
                background-color: #6C7A89;  /* ボタンの背景色 */
                color: white;  /* ボタンのテキストカラー */
                padding: 10px;
                border: 2px solid #B0B0B0;  /* ボーダーカラー */
                border-radius: 6px;
                min-width: 100px;  /* ボタンの最小幅 */
            }
            QPushButton:hover {
                background-color: #95A5A6;  /* ホバー時の背景色 */
            }
        """)

        # フォントを変更
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(9)
        self.setFont(font)

        # メッセージボックスのサイズを変更
        self.setFixedSize(2000, 300)

        # YesとNoボタンを水平に配置
        yes_button = self.addButton('Yes', QMessageBox.YesRole)
        no_button = self.addButton('No', QMessageBox.NoRole)

        # ボタンが存在する場合にスタイルシートを設定
        if yes_button is not None:
            yes_button.setStyleSheet("margin-right:30px;")  # 右マージン
        if no_button is not None:
            no_button.setStyleSheet("margin-left:30px;")  # 左マージン

def shutdown_pc():
    subprocess.run(["C:\Windows\system32\shutdown.exe", "/s", "/f", "/t", "2"])

def open_teams(user_directory):
    teams_path = fr"C:\Users\{user_directory}\AppData\Local\Microsoft\Teams\Update.exe --processStart Teams.exe"
    subprocess.run(teams_path)

def show_custom_message():
    app = QApplication([])
    dialog = CustomDialog()
    result = dialog.exec_()

    # StandardButtonを使って結果を確認
    if dialog.clickedButton() is not None:
        if dialog.clickedButton().text() == 'Yes':
            shutdown_pc()
        elif dialog.clickedButton().text() == 'No':
            user_directory = os.path.basename(os.path.expanduser("~"))
            open_teams(user_directory)

# カスタムメッセージボックスを表示
show_custom_message()
