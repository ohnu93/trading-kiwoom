import sys

from PyQt5.QtWidgets import QApplication

from autokiwoom.kiwoom import *

class Main():
    def __init__(self):
        print("main")

        self.app = QApplication(sys.argv)       # QApplication 객체 생성
        self.kiwoom = Kiwoom()
        self.app.exec_()                        # 이벤트 루프 실행

if __name__ == "__main__":
    Main()