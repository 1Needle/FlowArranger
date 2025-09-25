import json
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import QTimer

class Controller:
    def __init__(self):
        self.staff_table = None
        self.flow_table = None
        self.performance_table = None

        self.first_half = {}
        self.second_half = {}
        self.staff_dic = {}
        self.job_arr = [[]]
        self.unique_id = 0

        self.file_path = None

        self.auto_save_timer = QTimer()
        self.auto_save_timer.timeout.connect(lambda: self.save_file('auto'))
        self.auto_save_timer.start(60000)

    def clear_highlight(self):
        try:
            self.staff_table.clearHighlight()
            self.flow_table.clearHighlight()
            self.performance_table.clearHighlight()
        except Exception as e:
            print('Controller.py: clear_highlight', e)

    def read_file(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                None, "選擇檔案", "", "JSON Files (*.json);;All Files (*)"
            )
            if file_path:
                self.file_path = file_path
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                self.first_half = data["first_half"]
                self.second_half = data["second_half"]
                self.staff_dic = data["staff_dic"]
                self.job_arr = data["job_arr"]
                self.unique_id = data['unique_id']
                self.staff_table.update()
                self.flow_table.update()
                self.performance_table.update()
        except Exception as e:
            print('Controller.py: read_file', e)

    def save_file(self, mode):
        try:
            if mode == 'auto':
                file_path = 'auto_save.json'
            elif (mode == 'save' and not self.file_path) or mode == 'save as':
                file_path, _ = QFileDialog.getSaveFileName(
                    None, "儲存檔案", "save_file.json", "JSON Files (*.json);;All Files (*)"
                )
                if file_path:
                    self.file_path = file_path
            else:
                file_path = self.file_path

            if file_path:
                data = {
                    "first_half": self.first_half,
                    "second_half": self.second_half,
                    "staff_dic": self.staff_dic,
                    "job_arr": self.job_arr,
                    'unique_id': self.unique_id
                }

                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print('Controller.py: save_file', e)