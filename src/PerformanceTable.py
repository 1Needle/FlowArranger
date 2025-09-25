from PyQt5.QtWidgets import (QTableWidget, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                            QDialog, QFormLayout, QLineEdit, QSpinBox, QDialogButtonBox, QMessageBox,
                            QTableWidgetItem, QMenu, QAction)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QColor, QDrag

class CustomTable(QTableWidget):
    def __init__(self, parent, controller, row, col):
        try:
            super().__init__(row, col)
            self.setParent(parent)
            self.controller = controller
            self.setAcceptDrops(True)
            self.setDragEnabled(True)
            self.setDropIndicatorShown(True)
            self.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
            self.verticalHeader().customContextMenuRequested.connect(self.showRowMenu)
            self.setStyleSheet("""
                QTableWidget::item:selected {
                    background-color: #2828ff;
                }
            """)
        except Exception as e:
            print('PerformanceTable.py: __init__', e)

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
            print('PerformanceTable.py: showRowMenu', e)
        
    def deleteSelectedRows(self):
        try:
            selected_rows = set(index.row() for index in self.selectedIndexes())
            staff_dic = self.controller.staff_dic
            for row in sorted(selected_rows, reverse=True):
                p_name = self.item(row, 0).text()
                performers = self.item(row, 1).text().split(' ')
                assistants = self.item(row, 2).text().split(' ')
                for p in performers:
                    if p in staff_dic:
                        staff_dic[p]['performances'].remove(p_name)
                for a in assistants:
                    if a in staff_dic:
                        staff_dic[a]['assistances'].remove(p_name)
                for s in performers + assistants:
                    if s in staff_dic and not staff_dic[s]['performances'] and not staff_dic[s]['assistances']:
                        staff_dic.pop(s)
                self.removeRow(row)
            self.controller.staff_table.update()
            self.parent().updateFromTable()
        except Exception as e:
            print('PerformanceTable.py: deleteSelectedRows', e)

    def mousePressEvent(self, event):
        try:
            if event.button() == Qt.LeftButton:
                item = self.itemAt(event.pos())
                if item and item.column() == 0:
                    drag_row = item.row()
                    drag = QDrag(self)
                    mime = QMimeData()
                    mime.setText(str(drag_row))  # 傳遞原始 row 編號
                    drag.setMimeData(mime)
                    drag.exec_(Qt.MoveAction)
            super().mousePressEvent(event)
        except Exception as e:
            print('PerformanceTable.py: mousePressEvent', e)

    def dragEnterEvent(self, event):
        event.accept()

    def dragMoveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        """當第0欄被拖曳時，交換row"""
        try:
            source = event.source()
            if source == self.parent().first_half or source == self.parent().second_half:
                if event.mimeData().text():
                    source_row = int(event.mimeData().text())
                    pos = event.pos()
                    target_row = self.itemAt(pos).row()
                    self.parent().systemChange = True
                    for c in range(self.columnCount()):
                        temp = self.item(target_row, c).text()
                        self.setItem(target_row, c, QTableWidgetItem(source.item(source_row, c).text()))
                        source.setItem(source_row, c, QTableWidgetItem(temp))
                    self.parent().systemChange = False
                    self.parent().updateFromTable()
        except Exception as e:
            print('PerformanceTable.py: dropEvent', e)

class PerformanceTable(QWidget):
    def __init__(self, controller):
        try:
            super().__init__()

            self.systemChange = False
            self.controller = controller
            self.controller.performance_table = self

            # Table
            self.first_half = CustomTable(self, controller, 0, 4)
            self.first_half.setHorizontalHeaderLabels(['表演名稱', '表演人員', '場協', '表演時長'])
            self.first_half.cellPressed.connect(lambda row, col: self.pressed(row, col, 'first'))
            self.first_half.cellChanged.connect(lambda row, col: self.changed(row, col, 'first'))

            self.second_half = CustomTable(self, controller, 0, 4)
            self.second_half.setHorizontalHeaderLabels(['表演名稱', '表演人員', '場協', '表演時長'])
            self.second_half.cellPressed.connect(lambda row, col: self.pressed(row, col, 'second'))
            self.second_half.cellChanged.connect(lambda row, col: self.changed(row, col, 'second'))

            self.highlighted = []

            # Title & Button
            title1 = QLabel("上半場")
            title1.setAlignment(Qt.AlignCenter)
            title1.setStyleSheet("""
                font-size: 16px;
                font-weight: bold;
                font-family: Arial;
                color: white;
                background-color: #005AB5;
                border: 2px solid #003060;
                border-radius: 5px;
            """)
            title2 = QLabel("下半場")
            title2.setAlignment(Qt.AlignCenter)
            title2.setStyleSheet("""
                font-size: 16px;
                font-weight: bold;
                font-family: Arial;
                color: white;
                background-color: #005AB5;
                border: 2px solid #003060;
                border-radius: 5px;
            """)

            button1 = QPushButton("+")
            button1.setFixedSize(20, 20)
            button1.setToolTip("新增上半場表演")
            button1.clicked.connect(lambda: self.addPerformance('first'))
            button2 = QPushButton("+")
            button2.setFixedSize(20, 20)
            button2.setToolTip("新增下半場表演")
            button2.clicked.connect(lambda: self.addPerformance('second'))

            # SpinBox
            self.default_time = 5
            self.time = QSpinBox()
            self.time.setRange(0, 100)
            self.time.setValue(5)
            self.time.setToolTip("設定每支表演的預設時間\n改變此選項會重置所有表演的時間")
            self.time.valueChanged.connect(self.changeDefaultTime)

            # Sub Layout
            form_layout = QFormLayout()
            form_layout.addRow("預設表演時長:", self.time)

            top_layout = QHBoxLayout()
            top_layout.addWidget(title1)
            top_layout.addWidget(button1)

            middle_layout = QHBoxLayout()
            middle_layout.addWidget(title2)
            middle_layout.addWidget(button2)
            
            # Layout
            main_layout = QVBoxLayout()
            main_layout.addLayout(form_layout)
            main_layout.addLayout(top_layout)
            main_layout.addWidget(self.first_half)
            main_layout.addLayout(middle_layout)
            main_layout.addWidget(self.second_half)
            self.setLayout(main_layout)
        except Exception as e:
            print('PerformanceTable.py: __init___', e)

    def pressed(self, row, col, half):
        """點選表演時，以顏色表示所有工作人員在該場表演時是否有空"""
        try:
            self.controller.clear_highlight()
            h = None
            if half == 'first':
                h = self.first_half
                new_row = row + 1
            elif half == 'second':
                h = self.second_half
                new_row = row + len(self.controller.first_half) + 3
            item = h.item(row, col)
            if not item:
                return
            self.controller.staff_table.highlightPerformance(new_row)
        except Exception as e:
            print('PerformanceTable.py: pressed', e)

    def highlightPerformer(self, name):
        """給予表演者名稱，高亮該表演者對應的表演及場協"""
        try:
            self.systemChange = True
            staff_dic = self.controller.staff_dic
            if name not in staff_dic:
                return
            for p in staff_dic[name]['performances']: # 表演設為黃色
                found = False
                for row in range(self.first_half.rowCount()):
                    if p == self.first_half.item(row, 0).text():
                        self.first_half.item(row, 0).setBackground(QColor("#ffffaa"))
                        self.highlighted.append((row, 0, 'first'))
                        found = True
                        break
                if found:
                    continue
                for row in range(self.second_half.rowCount()):
                    if p == self.second_half.item(row, 0).text():
                        self.second_half.item(row, 0).setBackground(QColor("#ffffaa"))
                        self.highlighted.append((row, 0, 'second'))
                        break
            for p in staff_dic[name]['assistances']: # 場協設為橘色
                found = False
                for row in range(self.first_half.rowCount()):
                    if p == self.first_half.item(row, 0).text():
                        self.first_half.item(row, 0).setBackground(QColor("#ffaf60"))
                        self.highlighted.append((row, 0, 'first'))
                        found = True
                        break
                if found:
                    continue
                for row in range(self.second_half.rowCount()):
                    if p == self.second_half.item(row, 0).text():
                        self.second_half.item(row, 0).setBackground(QColor("#ffaf60"))
                        self.highlighted.append((row, 0, 'second'))
                        break
            self.systemChange = False
        except Exception as e:
            print('PerformanceTable.py: highlight performer', e)

    def highlightPerformance(self, row):
        """給予表演順序，高亮該表演"""
        try:
            self.systemChange = True
            if row == 0 or row == len(self.controller.first_half)+1 or row == len(self.controller.first_half)+2:
                return
            if row <= self.first_half.rowCount():
                row -= 1
                self.first_half.item(row, 0).setBackground(QColor("#ffffaa"))
                self.highlighted.append((row, 0, 'first'))
            else:
                row -= self.first_half.rowCount()+3
                self.second_half.item(row, 0).setBackground(QColor("#ffffaa"))
                self.highlighted.append((row, 0, 'second'))
            self.systemChange = False
        except Exception as e:
            print('PerformanceTable.py: highlight performance', e)

        
    def clearHighlight(self):
        """清除上次高亮顯示的物件"""
        try:
            self.systemChange = True
            for coor in self.highlighted:
                if coor[2] == 'first':
                    self.first_half.item(coor[0], coor[1]).setBackground(QColor('white'))
                if coor[2] == 'second':
                    self.second_half.item(coor[0], coor[1]).setBackground(QColor('white'))
            self.systemChange = False
        except Exception as e:
            print('PerformanceTable.py: clearHighlight', e)
        
    def addPerformance(self, half):
        """
        叫出輸入表演資訊的視窗，
        必填項沒有填寫的話會跳出錯誤視窗，
        根據[half]決定資料填入上半場或下半場
        """
        try:
            dialog = PerformanceInputDialog(default_time=self.default_time)
            if dialog.exec() == QDialog.Accepted:
                values = dialog.get_values()
                if half == "first":
                    dic = self.controller.first_half
                elif half == "second":
                    dic = self.controller.second_half

                dic[values['name']] = {'performers': values['performers'], 'assistants': values['assistants'], 'time': values['time']}
                self.controller.staff_table.addStaffList(values['performers'], values['name'], 'performer')
                self.controller.staff_table.addStaffList(values['assistants'], values['name'], 'assistant')
                self.update()

        except Exception as e:
            print('PerformanceTable.py: addPerformance', e)

    def changeDefaultTime(self):
        try:
            self.default_time = self.time.value()
        except Exception as e:
            print('PerformanceTable.py: changeDefaultTime', e)

    def numberPerformance(self):
        try:
            i = 1
            for name, dic in self.controller.first_half.items():
                dic['num'] = i
                i += 1
            i += 1
            for name, dic in self.controller.second_half.items():
                dic['num'] = i
                i += 1
        except Exception as e:
            print('PerformanceTable.py: numberPerformance', e)

    def update(self):
        """
        根據controller.first_half和controller.second_half中的資料
        將表演資料重新整理
        """
        try:
            self.systemChange = True
            self.numberPerformance()
            t = self.first_half
            t.setRowCount(0)
            row = 0
            for name, performance in self.controller.first_half.items():
                t.insertRow(row)
                t.setItem(row, 0, QTableWidgetItem(name))
                t.setItem(row, 1, QTableWidgetItem(" ".join(performance["performers"])))
                t.setItem(row, 2, QTableWidgetItem(" ".join(performance['assistants'])))
                t.setItem(row, 3, QTableWidgetItem(str(performance["time"])))
                row += 1
            t = self.second_half
            t.setRowCount(0)
            row = 0
            for name, performance in self.controller.second_half.items():
                t.insertRow(row)
                t.setItem(row, 0, QTableWidgetItem(name))
                t.setItem(row, 1, QTableWidgetItem(" ".join(performance["performers"])))
                t.setItem(row, 2, QTableWidgetItem(" ".join(performance['assistants'])))
                t.setItem(row, 3, QTableWidgetItem(str(performance["time"])))
                row += 1
            self.systemChange = False
        except Exception as e:
            print("PerformanceTable.py: update", e)

    def updateFromTable(self):
        """
        根據table中的資料
        重新整理controller.first_half和controller.second_half
        """
        def fillIn(t):
            new_dic = {}
            for row in range(t.rowCount()):
                name = t.item(row, 0).text()
                performers = t.item(row, 1).text().split(' ')
                assistants = t.item(row, 2).text().split(' ')
                time = int(t.item(row, 3).text())
                new_dic[name] = {'performers': performers, 'assistants': assistants, 'time': time}
            return new_dic
        
        try:
            self.controller.first_half = fillIn(self.first_half)
            self.controller.second_half = fillIn(self.second_half)
            self.numberPerformance()
            
        except Exception as e:
            print('PerformanceTable.py: updateFromTable', e)

    def changed(self, row, col, half_str):
        """若使用者手動更改表格內容，更新對應資料"""
        try:
            if self.systemChange == False:
                table = None
                half = None
                num = 0
                if half_str == 'first':
                    table = self.first_half
                    half = self.controller.first_half
                    num = row + 1
                elif half_str == 'second':
                    table = self.second_half
                    half = self.controller.second_half
                    num = row + len(self.controller.first_half) + 2
                name = table.item(row, 0).text()
                staff_table = self.controller.staff_table
                match col:
                    case 0:
                        for old_name, dic in half.items():
                            if dic['num'] == num:
                                staff_table.removeStaffList(half[old_name]['performers'], old_name, 'performer')
                                staff_table.removeStaffList(half[old_name]['assistants'], old_name, 'assistant')
                                break
                        half[name] = half.pop(old_name)
                        staff_table.addStaffList(half[name]['performers'], name, 'performer')
                        staff_table.addStaffList(half[name]['assistants'], name, 'assistant')
                    case 1:
                        performers = table.item(row, col).text().split(' ')
                        staff_table.removeStaffList(half[name]['performers'], name, 'performer')
                        staff_table.addStaffList(performers, name, 'performer')
                    case 2:
                        assistants = table.item(row, col).text().split(' ')
                        staff_table.removeStaffList(half[name]['assistants'], name, 'assistant')
                        staff_table.addStaffList(assistants, name, 'assistant')
                    case 3:
                        time = int(table.item(row, col).text())
                        half[name]['time'] = time
        except Exception as e:
            print('PerformanceTable.py: changed', e)


class PerformanceInputDialog(QDialog):
    def __init__(self, parent=None, default_time=None):
        super().__init__()

        self.setWindowTitle("新增表演")
        
        # Layout
        main_layout = QVBoxLayout()
        sub_layout = QHBoxLayout()
        form_layout1 = QFormLayout()
        form_layout2 = QFormLayout()

        # LineEdit
        self.performance_name_input = QLineEdit(self)
        self.performance_name_input.setFixedWidth(150)
        form_layout1.addRow("表演名稱(*必填):", self.performance_name_input)
        self.performance_people_input = QLineEdit(self)
        self.performance_people_input.setFixedWidth(300)
        form_layout1.addRow("表演人員(*必填):\n*請以空格隔開\nEx: 人員1 人員2 人員3", self.performance_people_input)
        self.performance_assistant_input = QLineEdit(self)
        self.performance_assistant_input.setFixedWidth(300)
        form_layout1.addRow("場協:", self.performance_assistant_input)

        # SpinBox
        self.performance_duration_input = QSpinBox(self)
        self.performance_duration_input.setRange(0, 100)
        self.performance_duration_input.setValue(default_time)
        form_layout2.addRow("表演時長:", self.performance_duration_input)


        # Adjust Layout
        form_layout1.setSpacing(20)
        sub_layout.addLayout(form_layout1)
        sub_layout.addLayout(form_layout2)
        main_layout.addLayout(sub_layout)

        # Add Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.check)
        button_box.rejected.connect(self.reject)
        button_box.setContentsMargins(20, 20, 20, 20)
        main_layout.addWidget(button_box)
        main_layout.setContentsMargins(30, 30, 30, 0)

        self.setLayout(main_layout)

    def check(self):
        """檢查必填項目，若未填寫則跳出錯誤視窗"""
        try:
            if not self.performance_name_input.text():
                QMessageBox.warning(self, "缺少輸入", "請輸入表演名稱!", QMessageBox.Ok)
            elif not self.performance_people_input.text():
                QMessageBox.warning(self, "缺少輸入", "請輸入表演人員!", QMessageBox.Ok)
            else:
                self.accept()
        except Exception as e:
            print('PerformanceTable.py: check', e)

    def get_values(self):
        """返回輸入的值"""
        try:
            return {
                'name': self.performance_name_input.text(),
                'time': self.performance_duration_input.value(),
                'performers': self.performance_people_input.text().split(' '),
                'assistants': self.performance_assistant_input.text().split(' ')
            }
        except Exception as e:
            print('PerformanceTable.py: get_values', e)