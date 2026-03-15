import sys
import time
import pandas as pd
import pyautogui
import pyperclip

from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog,
    QVBoxLayout, QLabel
)

class ExcelAutoInput(QWidget):

    def __init__(self):
        super().__init__()

        self.file_path = None

        self.setWindowTitle("Excel Auto Input")
        self.setGeometry(200,200,300,150)

        layout = QVBoxLayout()

        self.label = QLabel("엑셀 파일을 선택하세요")
        layout.addWidget(self.label)

        self.btn_file = QPushButton("엑셀 업로드")
        self.btn_file.clicked.connect(self.open_file)
        layout.addWidget(self.btn_file)

        self.btn_start = QPushButton("입력 시작")
        self.btn_start.clicked.connect(self.start_input)
        layout.addWidget(self.btn_start)

        self.setLayout(layout)


    def open_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls)"
        )

        if file_name:
            self.file_path = file_name
            self.label.setText(file_name)


    def start_input(self):

        if not self.file_path:
            self.label.setText("파일을 먼저 선택하세요")
            return

        df = pd.read_excel(self.file_path)

        # B열만 읽기
        values = df.iloc[:,1].dropna().tolist()

        self.label.setText("3초 후 시작합니다...")
        QApplication.processEvents()

        time.sleep(3)

        for v in values:

            text = str(v)

            pyperclip.copy(text)

            pyautogui.hotkey('ctrl','v')
            time.sleep(0.3)

            pyautogui.hotkey('alt','down')
            time.sleep(0.3)

        self.label.setText("완료")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = ExcelAutoInput()
    window.show()

    sys.exit(app.exec_())
