import os
import subprocess
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt


class CustomDialog(QMessageBox):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("お疲れ様でした")
        self.setText("終業時作業を忘れていませんか？")
        self.setInformativeText("Yes : シャットダウン実行\nNo  : シャットダウン中止. Teams を起動 ")
        self.setIcon(QMessageBox.Question)

        # スタイルシートを適用
        self.setStyleSheet("""
            QMessageBox {
                background-color: #333333;  /* Dark background color */
                color: #FFFFFF;  /* White text color */
                border: 2px solid #555555;  /* Dark border color */
                border-radius: 10px;
            }
            QLabel {
                color: #FFFFFF;  /* White text color */
                padding: 10px;
                margin: 0;
                font-size: 12pt;
                text-align: center;  /* Center-align the message */
            }
            QPushButton {
                background-color: #546E7A;  /* Dark button background color */
                color: white;  /* White button text color */
                padding: 10px;
                border: 2px solid #555555;  /* Dark border color */
                border-radius: 6px;
                min-width: 100px;  /* Minimum button width */
            }
            QPushButton:hover {
                background-color: #78909C;  /* Hover state background color */
                font-size: 10pt;
            }
        """)

        # フォントを変更
        font = QFont()
        font.setFamily("Meiryo")
        font.setPointSize(9)
        self.setFont(font)

        # メッセージボックスのサイズを変更
        self.setFixedSize(2000, 300)

        # YesとNoボタンを水平に配置
        yes_button = self.addButton('Yes', QMessageBox.YesRole)
        no_button = self.addButton('No', QMessageBox.NoRole)

        # ボタンが存在する場合にスタイルシートを設定
        if yes_button is not None:
            yes_button.setStyleSheet("margin-right: 80px; margin-left: 15px;")  # 左マージンを15pxに変更
            yes_button.setFont(QFont("Meiryo", 10))  # フォントサイズを12ptに変更

        if no_button is not None:
            no_button.setStyleSheet("margin-left: 0; margin-right: 15px;")  # 右マージンを15pxに変更
            no_button.setFont(QFont("Meiryo", 10))  # フォントサイズを12ptに変更

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
