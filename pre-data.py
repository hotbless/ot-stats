import openpyxl
import pyexcel
import datetime
import chinese_calendar



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
        return records


class WorkDay:
    def __init__(self):
        pass

    def is_workday(self):
        april_last = datetime.date(2020, 7, 25)
        result = chinese_calendar.is_workday(april_last)
        on_holiday, holiday_name = chinese_calendar.get_holiday_detail(april_last)
        print(result)
        print(on_holiday)
        print(holiday_name)






WorkDay().is_workday()
# ReadData().read_xlsx()


