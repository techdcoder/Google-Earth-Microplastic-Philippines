from openpyxl import Workbook


class ExcelInterface:
    def __init__(self, filename):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.filename = filename

    def generate(self, generator_function):
        generator_function(self.workbook)

    def read(self, read_function):
        return read_function(self.workbook)

    def save(self):
        self.workbook.save(self.filename)