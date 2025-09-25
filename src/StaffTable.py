from PyQt5.QtWidgets import (QTableWidget, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                            QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QMessageBox,
                            QTableWidgetItem, QMenu, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

class CustomTable(QTableWidget):
    def __init__(self, parent, controller, row, col):
        try:
            super().__init__(row, col)
            self.setParent(parent)
            self.controller = controller
            self.setAcceptDrops(False)
            self.setDragEnabled(True)
            self.setSortingEnabled(True)
            self.setDropIndicatorShown(True)
            self.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
            self.verticalHeader().customContextMenuRequested.connect(self.showRowMenu)
            self.setStyleSheet("""
                QTableWidget::item:selected {
                    background-color: #2828ff;
                }
            """)
        except Exception as e:
            print('StaffTable.py: __init__', e)

    def showRowMenu(self, pos):
        try:
            header = self.verticalHeader()
            row = header.logicalIndexAt(pos)
            if row != -1:
                menu = QMenu(self)
                delete_action = QAction("刪除列")
                delete_action.triggered.connect(self.deleteSelectedRows)
                menu.addAction(delete_action)
                menu.exec_(header.mapToGlobal(pos))
        except Exception as e:
            print('StaffTable.py: showRowMenu', e)

    def deleteSelectedRows(self):
        try:
            selected_rows = set(index.row() for index in self.selectedIndexes())
            staff_dic = self.controller.staff_dic
            for row in sorted(selected_rows, reverse=True):
                name = self.item(row, 0).text()
                staff_dic.pop(name)
                self.removeRow(row)
        except Exception as e:
            print('StaffTable.py: deleteSelectedRows', e)

class StaffTable(QWidget):
    def __init__(self, controller):
        try:
            super().__init__()

            self.controller = controller
            self.controller.staff_table = self
            self.systemChange = False

            # Table
            self.table = CustomTable(self, controller, 0, 7)
            self.table.setHorizontalHeaderLabels(['名字', '工作數量', '優先度', '火區', '表演', '場協', '編號'])
            self.table.setColumnHidden(6, True)
            self.table.cellClicked.connect(self.clicked)
            self.table.cellChanged.connect(self.changed)
            self.highlighted = []

            # Title & Button
            title = QLabel("工作人員")
            title.setAlignment(Qt.AlignCenter)
            title.setStyleSheet("""
                font-size: 16px;
                font-weight: bold;
                font-family: Arial;
                color: white;
                background-color: #005AB5;
                border: 2px solid #003060;
                border-radius: 5px;
            """)

            button = QPushButton("+")
            button.setFixedSize(20, 20)
            button.clicked.connect(self.addStaff)

            # Sub Layout
            top_layout = QHBoxLayout()
            top_layout.addWidget(title)
            top_layout.addWidget(button)

            # toolbar
            prio_label = QLabel("優先級:")
            prio_none_btn = QPushButton("無")
            prio_low_btn = QPushButton("低")
            prio_mid_btn = QPushButton("中")
            prio_high_btn = QPushButton("高")
            extinguish_label = QLabel('火區:')
            extinguish_yes_btn = QPushButton('可')
            extinguish_no_btn = QPushButton('不可')

            prio_none_btn.setToolTip("改變工作優先級 - 無工作")
            prio_low_btn.setToolTip("改變工作優先級 - 少數量工作")
            prio_mid_btn.setToolTip("改變工作優先級 - 正常數量工作")
            prio_high_btn.setToolTip("改變工作優先級 - 大量工作")
            extinguish_yes_btn.setToolTip("改變火區許可-可負責火區")
            extinguish_no_btn.setToolTip("改變火區許可-需有老人陪同")

            prio_none_btn.clicked.connect(lambda: self.changePrio('無'))
            prio_low_btn.clicked.connect(lambda: self.changePrio('低'))
            prio_mid_btn.clicked.connect(lambda: self.changePrio('中'))
            prio_high_btn.clicked.connect(lambda: self.changePrio('高'))
            extinguish_yes_btn.clicked.connect(lambda: self.changeExtinguish('可'))
            extinguish_no_btn.clicked.connect(lambda: self.changeExtinguish('不可'))

            toolbar_layout = QHBoxLayout()
            toolbar_layout.addWidget(prio_label)
            toolbar_layout.addWidget(prio_high_btn)
            toolbar_layout.addWidget(prio_mid_btn)
            toolbar_layout.addWidget(prio_low_btn)
            toolbar_layout.addWidget(prio_none_btn)
            toolbar_layout.addStretch()
            toolbar_layout.addWidget(extinguish_label)
            toolbar_layout.addWidget(extinguish_yes_btn)
            toolbar_layout.addWidget(extinguish_no_btn)

            # legend
            legend_layout = QHBoxLayout()
            items = [
                ("#53ff53", "空閒"),
                ("#ffffaa", "工作人員"),
                ("#ffaf60", "剛下場"),
                ("#FF7575", "準備"),
                ("#ff0000", "表演者")
            ]

            for color, text in items:
                block = QLabel()
                block.setFixedSize(20, 20)
                block.setStyleSheet(f"background-color: {color}; border: 1px solid black;")
                
                label = QLabel(text)
                label.setStyleSheet("margin-right: 3px;")

                legend_layout.addWidget(block)
                legend_layout.addWidget(label)


            # Layout
            main_layout = QVBoxLayout()
            main_layout.addLayout(top_layout)
            main_layout.addLayout(toolbar_layout)
            main_layout.addLayout(legend_layout)
            main_layout.addWidget(self.table)
            self.setLayout(main_layout)
        except Exception as e:
            print('StaffTable.py: __init__', e)

    def changed(self, row, col):
        try:
            if self.systemChange:
                return
            staff_dic = self.controller.staff_dic
            name = self.table.item(row, 0).text()
            if name in staff_dic:
                return
            num = int(self.table.item(row, 6).text())
            for old_name, dic in staff_dic.items():
                if dic['num'] == num:
                    staff_dic[name] = staff_dic.pop(old_name)
                    break
        except Exception as e:
            print("StaffTable.py: changed", e)

    def clicked(self, row, col):
        try:
            self.controller.clear_highlight()
            item = self.table.item(row, 0)
            if not item:
                return
            name = item.text()
            self.controller.flow_table.highlightStaff(name)
            self.controller.performance_table.highlightPerformer(name)
        except Exception as e:
            print('clicked', e)

    def highlightStaff(self, name):
        try:
            self.systemChange = True
            for row in range(self.table.rowCount()):
                if self.table.item(row, 0).text() == name:
                    self.table.item(row, 0).setBackground(QColor("#ffffaa"))
                    self.highlighted.append((row, 0))
                    break
            self.systemChange = False

        except Exception as e:
            print('StaffTable.py: highlightStaff', e)

    def highlightPerformance(self, row):
        """
        給予表演順序，將所有工作人員於該場表演時的狀態以顏色顯示
        空閒-綠色，準備二-淡紅色，準備一&表演者-紅色，剛下場-橘色
        """
        try:
            self.systemChange = True
            staff_dic = self.controller.staff_dic
            fh = self.controller.first_half
            sh = self.controller.second_half
            a = 'available_f'
            idx = row
            if row == len(fh)+1:
                return
            if row > len(fh)+1:
                a = 'available_s'
                idx -= len(fh)+2
            for r in range(self.table.rowCount()):
                staff_name = self.table.item(r, 0).text()
                if 'available_f' not in staff_dic[staff_name] or 'available_s' not in staff_dic[staff_name]:
                    continue
                if staff_dic[staff_name][a][idx] == 10: # 將可用人員變為綠色
                    self.table.item(r, 0).setBackground(QColor("#53ff53"))
                    self.highlighted.append((r, 0))
                elif staff_dic[staff_name][a][idx] == 7: # 將 表演下一隻 變為橘色
                    self.table.item(r, 0).setBackground(QColor("#ffaf60"))
                    self.highlighted.append((r, 0))
                elif staff_dic[staff_name][a][idx] == 3: # 將 準備12 變為淡紅色
                    self.table.item(r, 0).setBackground(QColor("#ff7575"))
                    self.highlighted.append((r, 0))
                elif staff_dic[staff_name][a][idx] == 1: # 將 工作人員 變為黃色
                    self.table.item(r, 0).setBackground(QColor("#ffffaa"))
                    self.highlighted.append((r, 0))
                elif staff_dic[staff_name][a][idx] == 0: # 將 表演者 變為紅色
                    self.table.item(r, 0).setBackground(QColor("#ff0000"))
                    self.highlighted.append((r, 0))
            self.systemChange = False
        except Exception as e:
            print('StaffTable.py: highlightPerformance', e)

    def clearHighlight(self):
        """清除所有高亮的儲存格"""
        try:
            self.systemChange = True
            for coor in self.highlighted:
                self.table.item(coor[0], coor[1]).setBackground(QColor('white'))
            self.systemChange = False
        except Exception as e:
            print('StaffTable.py: clearHighlight', e)

    def changePrio(self, prio):
        """將選取的row中所有人員的優先度都改為prio"""
        try:
            self.systemChange = True
            for index in self.table.selectedIndexes():
                row = index.row()
                name = self.table.item(row, 0).text()
                self.controller.staff_dic[name]['priority'] = prio
            self.update()
            self.systemChange = False
        except Exception as e:
            print('StaffTable.py: changePrio', e)

    def changeExtinguish(self, permission):
        """將選取的row中所有人員的火區允許權改為permission"""
        try:
            self.systemChange = True
            for index in self.table.selectedIndexes():
                row = index.row()
                name = self.table.item(row, 0).text()
                self.controller.staff_dic[name]['extinguish'] = permission
            self.update()
            self.systemChange = False
        except Exception as e:
            print('StaffTable.py: changeExtinguish', e)
        
    def addStaff(self):
        """呼叫使用者輸入視窗，可新增工作人員"""
        try:
            dialog = StaffInputDialog(self.controller.staff_dic)
            if dialog.exec_() == QDialog.Accepted:
                name = dialog.get_staff()
                dic = self.controller.staff_dic
                dic[name] = {"jobs": 0, "priority": '中', "performances": [], "assistances": [], 'extinguish': '可', 'num': self.controller.unique_id}
                self.controller.unique_id += 1
                self.update()
        except Exception as e:
            print('StaffTable.py: addStaff', e)

    def addStaffList(self, staff_list, p_name, type):
        """
        輸入名單，並標為表演者或場協
        staff_list: [list] 工作人員名單
        p_name: [str] 表演名稱
        type: [str] 'performer' or 'assistant'選擇標記
        """
        try:
            dic = self.controller.staff_dic
            for name in staff_list:
                if name:
                    if name not in dic:
                        dic[name] = {"jobs": 0, "priority": '中', "performances": [], "assistances": [], 'extinguish': '可', 'num': self.controller.unique_id}
                        self.controller.unique_id += 1
                    if type == 'performer':
                        dic[name]['performances'].append(p_name)
                    elif type == 'assistant':
                        dic[name]['assistances'].append(p_name)
            self.update()
        except Exception as e:
            print('StaffTable.py: addStaffList', e)

    def removeStaffList(self, staff_list, p_name, type):
        """
        輸入名單，並清除表演者或場協標記
        staff_list: [list] 工作人員名單
        p_name: [str] 表演名稱
        type: [str] 'performer' or 'assistant'選擇標記
        """
        try:
            dic = self.controller.staff_dic
            for name in staff_list:
                if not name:
                    continue
                if type == 'performer':
                        dic[name]['performances'].remove(p_name)
                elif type == 'assistant':
                        dic[name]['assistances'].remove(p_name)
                if not dic[name]['performances'] and not dic[name]['assistances']:
                    dic.pop(name)
            self.update()
        except Exception as e:
            print('StaffTable.py: removeStaffList', e)

    def update(self):
        """根據controller.staff_dic中的資料重新整理table"""
        try:
            self.systemChange = True
            self.table.setSortingEnabled(False)
            self.table.setRowCount(0)
            row = 0
            for name, dic in self.controller.staff_dic.items():
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(name))
                self.table.setItem(row, 1, QTableWidgetItem(str(dic["jobs"])))
                self.table.setItem(row, 2, QTableWidgetItem(dic["priority"]))
                self.table.setItem(row, 3, QTableWidgetItem(" ".join(dic["extinguish"])))
                self.table.setItem(row, 4, QTableWidgetItem(" ".join(dic["performances"])))
                self.table.setItem(row, 5, QTableWidgetItem(" ".join(dic["assistances"])))
                self.table.setItem(row, 6, QTableWidgetItem(str(dic["num"])))
                for i in range(1, 7):
                    item = self.table.item(row, i)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                row = row + 1
            self.table.setSortingEnabled(True)
            self.systemChange = False
        except Exception as e:
            print("StaffTable.py: update", e)

# 輸入視窗
class StaffInputDialog(QDialog):
    def __init__(self, staff_dic):
        try:
            super().__init__()
            self.staff_dic = staff_dic
            self.setWindowTitle("輸入工作人員")
            
            # 建立輸入欄位
            self.staff_input = QLineEdit()
            
            # 表單佈局
            form_layout = QFormLayout()
            form_layout.addRow("工作人員名字:", self.staff_input)

            # 確定/取消按鈕
            self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            self.button_box.accepted.connect(self.check)
            self.button_box.rejected.connect(self.reject)

            form_layout.addRow(self.button_box)
            self.setLayout(form_layout)
        except Exception as e:
            print("StaffTable.py: __init__", e)

    def check(self):
        """檢查是否未輸入，且名字是否已經存在"""
        try:
            name = self.staff_input.text()
            if not name:
                QMessageBox.warning(self, "缺少輸入", "請輸入工作人員名字!", QMessageBox.Ok)
                return
            if name in self.staff_dic:
                QMessageBox.warning(self, "重複輸入", "此工作人員已存在!", QMessageBox.Ok)
                return
            self.accept()
        except Exception as e:
            print("StaffTable.py: check", e)

    def get_staff(self):
        """回傳使用者輸入的名字"""
        return self.staff_input.text()