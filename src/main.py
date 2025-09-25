import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QAction, QVBoxLayout, QWidget, QSplitter)
from PyQt5.QtCore import Qt
from Controller import Controller
from StaffTable import StaffTable
from FlowTable import FlowTable
from PerformanceTable import PerformanceTable

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("細流編輯器")
        self.setGeometry(100, 100, 960, 540)
        
        self.initUI()
    # 
    def initUI(self):
        try:
            self.controller = Controller()

            # Toolbar
            menubar = self.menuBar()
            file_menu = menubar.addMenu("檔案")
            save_action = QAction("儲存檔案", self)
            save_as_action = QAction("另存新檔", self)
            read_action = QAction("讀取檔案", self)
            save_action.triggered.connect(lambda: self.controller.save_file('save'))
            save_as_action.triggered.connect(lambda: self.controller.save_file('save as'))
            read_action.triggered.connect(self.controller.read_file)
            file_menu.addAction(save_action)
            file_menu.addAction(save_as_action)
            file_menu.addAction(read_action)

            # Central Widget
            central_widget = QWidget()
            self.setCentralWidget(central_widget)

            # Tables
            self.left_table = PerformanceTable(self.controller)
            self.right_table = StaffTable(self.controller)
            self.middle_table = FlowTable(self.controller)

            # QSplitter
            self.splitter = QSplitter(Qt.Horizontal)
            self.splitter.setStyleSheet("""
                QSplitter::handle {
                    background-color: lightgray;
                    border: 1px solid black;
                    border-radius: 2px;
                    margin: 200px 3px
                }
            """)
            self.splitter.addWidget(self.left_table)
            self.splitter.addWidget(self.middle_table)
            self.splitter.addWidget(self.right_table)
            self.splitter.setSizes([250, 460, 250])

            # Main Layout
            main_layout = QVBoxLayout()
            main_layout.addWidget(self.splitter)
            central_widget.setLayout(main_layout)
        except Exception as e:
            print('main.py: initUI', e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # apply_stylesheet(app, theme='dark_blue.xml', invert_secondary=True)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())