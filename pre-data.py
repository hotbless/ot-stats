import openpyxl
import pyexcel



class ReadData:
    def __init__(self):
        pass
        self.filename = "20200717.xlsx"

    def read_data(self):
        wb = openpyxl.load_workbook(filename=self.filename)
        print(wb)

    def read_xlsx(self):
        records = pyexcel.get_records(file_name=self.filename)
        print(records[0])
        print(type(records[0]))
        print(records[0].keys())



ReadData().read_xlsx()


