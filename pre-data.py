import openpyxl
import pyexcel
import datetime
import chinese_calendar



class ReadOtData:
    def __init__(self):
        pass
        self.filename = "20200717.xlsx"

    def read_data(self):
        wb = openpyxl.load_workbook(filename=self.filename)
        print(wb)

    def read_xlsx(self):
        records = pyexcel.get_records(file_name=self.filename)
        # print(records[0])
        # print(type(records[0]))
        # print(records[0].keys())
        return records


class WorkDay:
    def __init__(self):
        pass

    def is_workday(self):
        # april_last = datetime.date(2020, 10, 1)
        str_p = '2020-06-20 09:40'
        april_last = datetime.datetime.strptime(str_p, '%Y-%m-%d %H:%M')
        print(april_last)

        result = chinese_calendar.is_workday(april_last)
        on_holiday, holiday_name = chinese_calendar.get_holiday_detail(april_last)
        print("is work day?", result)
        print("is holiday?", on_holiday)
        print("holiday name?", holiday_name)

# Check if the date is workday and return
class IsWorkDay:
    def __init__(self, w_date):
        # w_date_f = datetime.datetime.strptime(w_date, '%Y-%m-%d %H:%M')

        # self.w_res = chinese_calendar.is_workday(w_date_f)
        self.w_res = chinese_calendar.is_workday(w_date)
        on_holiday, holiday_name = chinese_calendar.get_holiday_detail(w_date)
        print("is work day?", self.w_res)
        print("is holiday?", on_holiday)
        print("holiday name?", holiday_name)

    def res(self):
        return self.w_res


class StrToDatetimeHM:
    def __init__(self, str_time):
        self.res = datetime.datetime.strptime(str_time, '%Y-%m-%d %H:%M')
        print(type(self.res))

    def res(self):
        return self.res


class OtHours:
    def __init__(self, ot_start, ot_end):
        print(type(ot_start))
        print(type(ot_end))
        self.ot_hours = (ot_end - ot_start).seconds
        print(self.ot_hours)

    def res(self):
        return self.ot_hours


class CalDate:
    def __init__(self):
        pass

    def calculate_date(self):
        ot_records = ReadOtData().read_xlsx()
        # ot_name = ot_records[0]['over']
        # print(ot_name)
        ot_start = StrToDatetimeHM(ot_records[0]['开始时间']).res
        print(type(ot_start))
        d_res = IsWorkDay(ot_start).res
        if d_res is True:
            print("workday")
        else:
            print("holiday")
        ot_end = StrToDatetimeHM(ot_records[0]['结束时间']).res
        ot_hours = OtHours(ot_start, ot_end).res
        print(ot_hours)
        # for ot_r in ot_records:
        #     print()



# if __name__ == '__main__':
WorkDay().is_workday()
CalDate().calculate_date()
    # ReadOtData().read_xlsx()


