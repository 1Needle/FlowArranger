from PyQt5.QtWidgets import QTableWidget, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidgetItem, QMessageBox, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from pandas import DataFrame
import openpyxl

class CustomTable(QTableWidget):
    def __init__(self, controller, row, col):
        try:
            super().__init__(row, col)
            self.controller = controller
            self.setAcceptDrops(True)
            self.setDragEnabled(True)
            self.setDropIndicatorShown(True)
            self.setStyleSheet("""
                QTableWidget::item:selected {
                    background-color: #2828ff;
                }
            """)
        except Exception as e:
            print('FlowTable.py: __init__', e)

    def dropEvent(self, event):
        """無視對表演名稱的更改，允許更改工作人員"""
        try:
            pos = event.pos()
            item = self.itemAt(pos)
            if event.source() == self.controller.staff_table.table:
                pos = event.pos()
                item = self.itemAt(pos)
                if item.column() < 3 or item.row() == len(self.controller.first_half)+1:
                    event.ignore()
                else:
                    super().dropEvent(event)
                    self.parent().highlighted.append((item.row(), item.column()))
            if event.source() == self:
                super().dropEvent(event)

        except Exception as e:
            print('FlowTable.py: dropEvent', e)

class FlowTable(QWidget):
    def __init__(self, controller):
        try:
            super().__init__()

            self.systemChange = False

            self.controller = controller
            self.controller.flow_table = self

            # Table
            self.table = CustomTable(self.controller, 0, 13)
            self.table.setHorizontalHeaderLabels(['表演者', '準備一', '準備二', '闈場1', '闈場2', '闈場3', '闈場4', '闈場5', '闈場6', '左火區1', '左火區2', '右火區1', '右火區2'])
            self.table.cellChanged.connect(self.changed)

            # Title & Button
            title = QLabel("細流")
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
            output_button = QPushButton("輸出表格")
            output_button.setFixedSize(80, 20)
            output_button.clicked.connect(self.output)

            check_button = QPushButton("自動檢查")
            check_button.setFixedSize(80, 20)
            check_button.clicked.connect(self.check)

            generate_button = QPushButton("生成細流")
            generate_button.setFixedSize(80, 20)
            generate_button.clicked.connect(self.generate)

            # Sub Layout
            top_layout = QHBoxLayout()
            top_layout.addWidget(title)
            top_layout.addWidget(output_button)
            top_layout.addWidget(check_button)
            top_layout.addWidget(generate_button)

            # Layout
            main_layout = QVBoxLayout()
            main_layout.addLayout(top_layout)
            main_layout.addWidget(self.table)
            self.setLayout(main_layout)
            
            self.highlighted = []
            self.table.cellClicked.connect(self.clicked)
        except Exception as e:
            print('FlowTable.py: __init__', e)

    def changed(self, row, col):
        """
        使用者改變格子內容時，更新對應資料
        """
        try:
            if self.systemChange == True:
                return
            if col < 3:
                return
            fh = self.controller.first_half
            sh = self.controller.second_half
            staff_dic = self.controller.staff_dic
            arr = self.controller.job_arr
            arr_row = row if row <= len(fh) else row-1
            prev_name = arr[arr_row][col-3]
            name = self.table.item(row, col).text()
            if name != prev_name:
                arr[arr_row][col-3] = name
                half = fh if row <= len(fh) else sh
                available = 'available_f' if row <= len(fh) else 'available_s'
                available_row = row if row <= len(fh) else row-1-len(fh)
                staff_dic[name]['jobs'] += 1
                staff_dic[name]['job_index'].append([row, col])
                staff_dic[name][available][available_row] = 1
                if prev_name:
                    staff_dic[prev_name]['jobs'] -= 1
                    staff_dic[prev_name]['job_index'].remove([row, col])

                    a = [10] * (len(half) + 1)
                    i = 1
                    for p in half.keys():
                        if p in staff_dic[prev_name]["performances"]:
                            a[0] = 0
                            a[i] = 0
                            if i-1 > 0:
                                a[i-1] = 3 # 降低準備一
                            if i-2 > 0:
                                a[i-2] = 3 # 降低準備二
                            if i+1 <= len(half): 
                                a[i+1] = 7 # 降低下一場
                        elif p in staff_dic[prev_name]['assistances']:
                            a[i] = 0
                            if i-1 > 0:
                                a[i-1] = 3
                        else:
                            for r, _ in staff_dic[prev_name]['job_index']:
                                if r == row:
                                    a[i] = 1
                        i += 1

                    staff_dic[prev_name][available] = a
                self.controller.staff_table.update()
        except Exception as e:
            print('FlowTable.py: changed', e)

    def clicked(self, row, col):
        """
        點擊工作人員->高亮其他工作、顯示無法工作的場次
        點擊表演名稱->高亮對應表演、顯示該表演時所有工作人員的狀態
        """
        try:
            self.controller.clear_highlight()
            item = self.table.item(row, col)
            if not item:
                return
            name = item.text()
            if col >= 3: # 若點擊工作人員
                self.highlightStaff(name)
                self.controller.staff_table.highlightStaff(name)
                self.controller.performance_table.highlightPerformer(name)
            elif col >= 0 and col < 3: # 若點擊表演名稱
                self.controller.performance_table.highlightPerformance(row)
                self.controller.staff_table.highlightPerformance(row)
        except Exception as e:
            print('FlowTable.py: clicked', e)

    def highlightStaff(self, name):
        """給予工作人員名字，高亮其在細流中的表演及工作"""
        try:
            self.systemChange = True
            staff_dic = self.controller.staff_dic
            fh = self.controller.first_half
            sh = self.controller.second_half
            if name not in staff_dic:
                return
            for p in staff_dic[name]['performances']: # 標出所有上半場表演
                if p in fh:
                    num = fh[p]['num']
                    for c in range(self.table.columnCount()):
                        self.table.item(num, c).setBackground(QColor("#FF0000"))
                        self.highlighted.append((num, c))
                        if num - 1 >= 0:
                            self.table.item(num-1, c).setBackground(QColor("#FF7575"))
                            self.highlighted.append((num-1, c))
                        if num - 2 >= 0:
                            self.table.item(num-2, c).setBackground(QColor("#FF7575"))
                            self.highlighted.append((num-2, c))
                        if num + 1 <= len(fh):
                            self.table.item(num+1, c).setBackground(QColor("#ffaf60"))
                            self.highlighted.append((num+1, c))
            for p in staff_dic[name]['performances']: # 標出所有下半場表演
                if p in sh:
                    num = sh[p]['num'] + 1
                    for c in range(self.table.columnCount()):
                        self.table.item(num, c).setBackground(QColor("#FF0000"))
                        self.highlighted.append((num, c))
                        if num - 1 > len(fh):
                            self.table.item(num-1, c).setBackground(QColor("#FF7575"))
                            self.highlighted.append((num-1, c))
                        if num - 2 > len(fh):
                            self.table.item(num-2, c).setBackground(QColor("#FF7575"))
                            self.highlighted.append((num-2, c))
                        if num + 1 < self.table.rowCount():
                            self.table.item(num+1, c).setBackground(QColor("#ffaf60"))
                            self.highlighted.append((num+1, c))
            if 'job_index' not in staff_dic[name]:
                return
            for coor in staff_dic[name]['job_index']: # 標出所有工作
                self.table.item(coor[0], coor[1]).setBackground(QColor("#ffffaa"))
                self.highlighted.append((coor[0], coor[1]))
            self.systemChange = False
        except Exception as e:
            print('FlowTable.py: highlightStaff', e)

    def clearHighlight(self):
        """清除高亮的儲存格"""
        try:
            self.systemChange = True
            for coor in self.highlighted:
                self.table.item(coor[0], coor[1]).setBackground(QColor('white'))
            self.systemChange = False
        except Exception as e:
            print('FlowTable.py: clearHighlight', e)
        
    def generate(self):
        """根據controller.first_half和controller.second_half和controller.staff_dic中的資料自動生成細流"""
        try:
            fh = self.controller.first_half
            sh = self.controller.second_half
            staff_dic = self.controller.staff_dic
            max_row = len(fh) + len(sh) + 2
            arr = [["" for _ in range(10)] for _ in range(max_row)]

            # ========= 一些初始化 =========
            self.controller.performance_table.numberPerformance()

            # ========= 計算各種權重 ========
            for p in fh.values():
                p['weight'] = 0
            for p in sh.values():
                p['weight'] = 0

            for staff in staff_dic.values():
                if staff['priority'] == '高':
                    staff['limit'] = 6
                elif staff['priority'] == '中':
                    staff['limit'] = 4
                elif staff['priority'] == '低':
                    staff['limit'] = 2
                elif staff['priority'] == '無':
                    staff['limit'] = 0
                staff['jobs'] = 0
                staff['job_index'] = []

                # 計算人員空閒時間
                available_f = [10] * (len(fh) + 1)
                i = 1
                for performance in fh.keys():
                    if performance in staff["performances"]:
                        available_f[0] = 0 # 禁止預熱闈場
                        available_f[i] = 0 # 禁止該表演
                        if i-1 > 0:
                            available_f[i-1] = 3 # 禁止準備一
                        if i-2 > 0:
                            available_f[i-2] = 3 # 降低準備二
                        if i+1 < len(fh) + 1: 
                            available_f[i+1] = 7 # 降低下一場
                    elif performance in staff["assistances"]:
                        available_f[i] = 0 # 禁止該表演
                        if i-1 > 0:
                            available_f[i-1] = 3 # 降低準備一
                    i += 1
                i = 1
                available_s = [10] * (len(sh) + 1)
                for performance in sh.keys():
                    if performance in staff["performances"]:
                        available_s[0] = 0 # 禁止預熱闈場
                        available_s[i] = 0 # 禁止該表演
                        if i-1 > 0:
                            available_s[i-1] = 3 # 禁止準備一
                        if i-2 > 0:
                            available_s[i-2] = 3 # 降低準備二
                        if i+1 < len(sh) + 1: 
                            available_s[i+1] = 7 # 降低下一場
                    elif performance in staff["assistances"]:
                        available_s[i] = 0 # 禁止該表演
                        if i-1 > 0:
                            available_s[i-1] = 3 # 降低準備一
                    i += 1
                    
                staff["available_f"] = available_f
                staff["available_s"] = available_s
            
                # 計算表演權重
                i = 1
                for p in fh.values():
                    p["weight"] += available_f[i]
                    i += 1
                i = 1
                for p in sh.values():
                    p["weight"] += available_s[i]
                    i += 1

            # 一次塞兩份工作 - 上半場
            copy_h = dict(fh)
            while copy_h:
                sorted_h = sorted(copy_h.items(), key=lambda item: item[1]["weight"])
                target = sorted_h[0]
                row = target[1]["num"]
                sorted_s = sorted(staff_dic.items(), key=lambda s: (s[1]['jobs']))

                col = 0
                offset = -1 if row % 2 == 0 else 1
                for s in sorted_s:
                    if row + offset < 0 or row + offset > len(fh): # 若offset超出範圍則跳出
                        break
                    if col == 10: # 若填完則跳出
                        break
                    if arr[row][col]: # 若已填入則跳過
                        col += 1
                        continue
                    if s[1]["jobs"] + 2 > s[1]['limit']: # 限制個人工作數量上限
                        continue
                    if (row - 2 > 0 and s[0] in arr[row - 2]) or (row + 2 < len(fh)+1 and s[0] in arr[row + 2]): # 防止連續3+場工作
                        continue
                    if s[1]["available_f"][row] == 10 and s[1]["available_f"][row+offset] == 10:
                        if (col == 6 or col == 8) and s[1]['extinguish'] == '不可':
                            continue
                        arr[row][col] = s[0]
                        arr[row+offset][col] = s[0]
                        s[1]["available_f"][row] = 1
                        s[1]["available_f"][row+offset] = 1
                        s[1]["job_index"].append([row, col+3])
                        s[1]["job_index"].append([row+offset, col+3])
                        s[1]["jobs"] += 2
                        col += 1
            
                copy_h.pop(target[0])
                for name, dic in copy_h.items():
                    if dic["num"] == row + offset:
                        copy_h.pop(name)
                        break

            # 一次塞兩個工作-下半場
            copy_h = dict(sh)
            while copy_h:
                sorted_h = sorted(copy_h.items(), key=lambda item: item[1]["weight"])
                target = sorted_h[0]
                row = target[1]["num"]
                idx = row - len(fh) - 1
                sorted_s = sorted(staff_dic.items(), key=lambda s: (-s[1]["available_s"][idx], s[1]['jobs']))

                col = 0
                offset = -1 if idx % 2 == 0 else 1
                for s in sorted_s:
                    if idx + offset < 0 or idx + offset > len(sh): # 若offset超出範圍則跳出
                        break
                    if col == 10: # 若填完則跳出
                        break
                    if arr[row][col]: # 若已填入則跳過
                        col += 1
                        continue
                    if s[1]["jobs"] + 2 > s[1]['limit']: # 限制個人工作數量上限
                        continue
                    if (row - 2 > 0 and s[0] in arr[row - 2]) or (row + 2 < max_row and s[0] in arr[row + 2]): # 防止連續3+場工作
                        continue
                    if s[1]["available_s"][idx] == 10 and s[1]["available_s"][idx+offset] == 10:
                        if (col == 6 or col == 8) and s[1]['extinguish'] == '不可':
                            continue
                        arr[row][col] = s[0]
                        arr[row+offset][col] = s[0]
                        s[1]["available_s"][idx] = 1
                        s[1]["available_s"][idx+offset] = 1
                        s[1]["job_index"].append([row+1, col+3])
                        s[1]["job_index"].append([row+offset+1, col+3])
                        s[1]["jobs"] += 2
                        col += 1
                
                copy_h.pop(target[0])
                for name, dic in copy_h.items():
                    if dic["num"] == row + offset:
                        copy_h.pop(name)
                        break
                
            # 補空位 - 上半場
            copy_h = dict(fh)
            while copy_h:
                sorted_h = sorted(copy_h.items(), key=lambda item: item[1]["weight"])
                target = sorted_h[0]
                row = target[1]["num"]
                sorted_s = sorted(staff_dic.items(), key=lambda s: (-s[1]["available_f"][row], s[1]['jobs']))

                col = 0
                offset = -1 if row % 2 == 0 else 1
                for s in sorted_s:
                    if col == 10: # 若填完則跳出
                        break
                    if arr[row][col]: # 若已填入則跳過
                        col += 1
                        continue
                    if s[1]["jobs"] + 1 > s[1]['limit']: # 限制個人工作數量上限
                        continue
                    if (row - 1 > 0 and s[0] in arr[row - 1]) or (row + 1 < len(fh)+1 and s[0] in arr[row + 1]): # 防止連續3+場工作
                        continue
                    if s[1]["available_f"][row] == 10:
                        if (col == 6 or col == 8) and s[1]['extinguish'] == '不可':
                            continue
                        arr[row][col] = s[0]
                        s[1]["available_f"][row] = 1
                        s[1]["job_index"].append([row, col+3])
                        s[1]["jobs"] += 1
                        col += 1
                
                copy_h.pop(target[0])

            # 補空位 - 下半場
            copy_h = dict(sh)
            while copy_h:
                sorted_h = sorted(copy_h.items(), key=lambda item: item[1]["weight"])
                target = sorted_h[0]
                row = target[1]["num"]
                idx = row - len(fh) - 1
                sorted_s = sorted(staff_dic.items(), key=lambda s: (-s[1]["available_s"][idx], s[1]['jobs']))

                col = 0
                offset = -1 if row % 2 == 0 else 1
                for s in sorted_s:
                    if col == 10: # 若填完則跳出
                        break
                    if arr[row][col]: # 若已填入則跳過
                        col += 1
                        continue
                    if s[1]["jobs"] + 1 > s[1]['limit']: # 限制個人工作數量上限
                        continue
                    if (row - 1 > 0 and s[0] in arr[row - 1]) or (row + 1 < max_row and s[0] in arr[row + 1]): # 防止連續3+場工作
                        continue
                    if s[1]["available_s"][idx] == 10:
                        if (col == 6 or col == 8) and s[1]['extinguish'] == '不可':
                            continue
                        arr[row][col] = s[0]
                        s[1]["available_s"][idx] = 1
                        s[1]["job_index"].append([row+1, col+3])
                        s[1]["jobs"] += 1
                        col += 1
                
                copy_h.pop(target[0])

            # 補預熱
            sorted_s = sorted(staff_dic.items(), key=lambda s: (-s[1]["available_f"][0], s[1]['jobs']))
            col = 0
            for s in sorted_s:
                if col == 10:
                    break
                if s[1]["jobs"] + 1 > s[1]['limit']: # 限制個人工作數量上限
                    continue
                if s[1]["available_f"][0] == 10:
                    if (col == 6 or col == 8) and s[1]['extinguish'] == '不可':
                        continue
                    arr[0][col] = s[0]
                    s[1]["available_f"][0] = 1
                    s[1]["job_index"].append([0, col+3])
                    s[1]["jobs"] += 1
                    col += 1

            sorted_s = sorted(staff_dic.items(), key=lambda s: (-s[1]["available_s"][0], s[1]['jobs']))
            col = 0
            for s in sorted_s:
                if col == 10:
                    break
                if s[1]["jobs"] + 1 > s[1]['limit']: # 限制個人工作數量上限
                    continue
                if s[1]["available_s"][0] == 10:
                    if (col == 6 or col == 8) and s[1]['extinguish'] == '不可':
                        continue
                    arr[len(fh)+1][col] = s[0]
                    s[1]["available_s"][0] = 1
                    s[1]["job_index"].append([len(fh)+2, col+3])
                    s[1]["jobs"] += 1
                    col += 1

            self.controller.job_arr = arr
            self.update()
            self.controller.staff_table.update()
        except Exception as e:
            print('FlowTable.py: generate', e)

    def update(self):
        """根據controller.job_arr的資料重新整理table"""
        try:
            self.systemChange = True
            # 初始化
            max_row = len(self.controller.first_half) + len(self.controller.second_half) + 2
            self.table.setRowCount(0)
            self.table.setRowCount(max_row+1)
            for r in range(self.table.rowCount()):
                for c in range(self.table.columnCount()):
                    self.table.setItem(r, c, QTableWidgetItem())

            # 填入表演順序
            arr = self.controller.job_arr
            fh = self.controller.first_half
            sh = self.controller.second_half
            self.table.setItem(0, 0, QTableWidgetItem("上半場預熱"))
            self.table.setItem(len(fh)+1, 0, QTableWidgetItem("中場休息"))
            self.table.setItem(len(fh)+2, 0, QTableWidgetItem("下半場預熱"))
            row = 1
            for name in fh.keys():
                self.table.setItem(row, 0, QTableWidgetItem(name))
                if row-1 > 0 :
                    self.table.setItem(row-1, 1, QTableWidgetItem(name))
                if row-2 > 0:
                    self.table.setItem(row-2, 2, QTableWidgetItem(name))
                row += 1
            row += 2
            for name in sh.keys():
                self.table.setItem(row, 0, QTableWidgetItem(name))
                if row-1 > len(fh)+2 :
                    self.table.setItem(row-1, 1, QTableWidgetItem(name))
                if row-2 > len(fh)+2:
                    self.table.setItem(row-2, 2, QTableWidgetItem(name))
                row += 1

            # 設定表演名稱唯讀
            for row in range(self.table.rowCount()):
                for col in range(3):
                    item = self.table.item(row, col)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            # 填入工作人員
            for row in range(len(arr)):
                for col in range(len(arr[0])):
                    new_r = row
                    if row >= len(fh)+1:
                        new_r = row + 1
                    self.table.setItem(new_r, col+3, QTableWidgetItem(arr[row][col]))
            self.systemChange = False
        except Exception as e:
            print("FlowTable.py: update", e)

    def check(self):
        try:
            staff_dic = self.controller.staff_dic
            p_list = []
            pre1_list = []
            pre2_list = []
            next_list = []
            assist_list = []
            pre1_a_list = []
            for row in range(self.table.rowCount()):
                p = self.table.item(row, 0).text()
                pre1 = self.table.item(row, 1).text()
                pre2 = self.table.item(row, 2).text()
                for col in range(3, self.table.columnCount()):
                    name = self.table.item(row, col).text()
                    if not name:
                        continue
                    performances = staff_dic[name]['performances']
                    assistances = staff_dic[name]['assistances']
                    if p in performances:
                        p_list.append(name)
                    if p in assistances:
                        assist_list.append(name)
                    if pre1 in performances:
                        pre1_list.append(name)
                    if pre1 in assistances:
                        pre1_a_list.append(name)
                    if pre2 in performances:
                        pre2_list.append(name)
                    if row-1 > 0:
                        next = self.table.item(row-1, 0).text()
                        if next in performances:
                            next_list.append(name)
            text = ''
            if p_list:
                text += '以下人員的 表演 與工作重疊了:\n'
                text += ', '.join(p_list) + '\n\n'
            if pre1_list:
                text += '以下人員的 準備1 與工作重疊了:\n'
                text += ', '.join(pre1_list) + '\n\n'
            if pre2_list:
                text += '以下人員的 準備2 與工作重疊了:\n'
                text += ', '.join(pre2_list) + '\n\n'
            if next_list:
                text += '以下人員的 表演完下一場 與工作重疊了:\n'
                text += ', '.join(next_list) + '\n\n'
            if assist_list:
                text += '以下人員的 場協 與工作重疊了:\n'
                text += ', '.join(assist_list) + '\n\n'
            if pre1_a_list:
                text += '以下人員的 場協準備1 與工作重疊了:\n'
                text += ', '.join(pre1_a_list) + '\n\n'
            if not text:
                text+= '沒有任何問題!'
            
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle('自動檢查完成')
            msg.setText(text)
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

        except Exception as e:
            print('FlowTable.py: check', e)

    def output(self):
        try:
            table_widget = self.table
            row_count = table_widget.rowCount()
            column_count = table_widget.columnCount()

            # 取得表頭
            headers = ['表演者', '準備一', '準備二', '闈場(1.2.3.4.5.6)', '左火區 / 右火區']

            # 收集表格資料
            data = []
            for row in range(row_count):
                if row == len(self.controller.first_half)+1:
                    data.append(['中場休息'])
                    continue
                row_data = []
                security = []
                left_extinguish = []
                right_extinguish = []
                for col in range(column_count):
                    item = table_widget.item(row, col)
                    if col < 3:
                        row_data.append(item.text() if item else "")
                    elif col < 9:
                        security.append(item.text() if item else "")
                    elif col < 11:
                        left_extinguish.append(item.text() if item else "")
                    else:
                        right_extinguish.append(item.text() if item else "")
                security = '、'.join(security)
                left_extinguish = '、'.join(left_extinguish)
                right_extinguish = '、'.join(right_extinguish)
                all_extinguish = left_extinguish + ' / ' + right_extinguish
                row_data.append(security)
                row_data.append(all_extinguish)
                data.append(row_data)

            # 建立 DataFrame 並匯出為 Excel
            df = DataFrame(data, columns=headers)

            # 選擇儲存檔案路徑
            file_path, _ = QFileDialog.getSaveFileName(None, "Save File", "output.xlsx", "Excel Files (*.xlsx)")
            if file_path:
                df.to_excel(file_path, index=False)
                print(f"Table exported to {file_path}")
        except Exception as e:
            print('FlowTable.py: output', e)