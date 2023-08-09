import openpyxl
class Openfile:
    def __init__(self):
        self.worksheet = None

    def open(self, file_path):
        self.file_path = file_path
        try:
            self.workbook = openpyxl.load_workbook(self.file_path)
            self.worksheet = self.workbook.active
        except Exception as e:
            print(f"Произошла ошибка при открытии файла: {e}")
    def get_worksheet(self):
        return self.worksheet