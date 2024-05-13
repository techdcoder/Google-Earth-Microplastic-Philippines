from openpyxl import Workbook


class ExcelInterface:
    def __init__(self, filename):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active()
        self.filename = filename

    def save(self):
        self.workbook.save(self.filename)