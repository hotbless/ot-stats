import openpyxl
import datetime
import chinese_calendar



# class StrToDatetime:
#     def __init__(self, str_time):
#         self.dt_day = str_time.split()[0]
#         self.dt_time = str_time.split()[1]
#
#     def dt_hrmin(self):
#         self.dt_time = datetime.datetime.strptime(self.dt_time, '%H:%M')
#
#     def dt_ymd(self):
#         self.dt_day = datetime.datetime.strptime(self.dt_day, '%Y-%m-%d')
#
#     @property
#     def res(self):
#         self.dt_ymd()
#         self.dt_hrmin()
#         return self.dt_day, self.dt_time


class StrToDatetime:
    def __init__(self, str_time):
        self.time_tr = str_time

    def dt_hrmin(self):
        hm = datetime.datetime.strptime(self.time_tr, '%H:%M')
        return hm

    def dt_ymd(self):
        ymd = datetime.datetime.strptime(self.time_tr, '%Y-%m-%d')
        return ymd


class ElasTime:
    def __init__(self, time_start, time_end):
        self.time_start = time_start
        self.time_end = time_end

    def hm_str_to_dt(self):
        self.time_start = StrToDatetime(self.time_start).dt_hrmin()
        self.time_end = StrToDatetime(self.time_end).dt_hrmin()

    def elas_hrs(self):
        self.hm_str_to_dt()
        es = (self.time_end - self.time_start).seconds/3600
        return es


class IsWorkDay:
    def __init__(self, w_date):
        # w_date_f = datetime.datetime.strptime(w_date, '%Y-%m-%d %H:%M')

        # self.w_res = chinese_calendar.is_workday(w_date_f)
        self.w_res = chinese_calendar.is_workday(w_date)
        on_holiday, holiday_name = chinese_calendar.get_holiday_detail(w_date)
        print("is work day?", self.w_res)
        print("is holiday?", on_holiday)
        print("holiday name?", holiday_name)

    @property
    def res(self):
        return self.w_res


class ReadOtData:
    def __init__(self):
        self.filename = "20200717.xlsx"
        self.wb = openpyxl.load_workbook(filename=self.filename)
        self.save_name = "new_ot.xlsx"

    def unmerge(self):
        for ws in self.wb:
            print(ws.title)
            merged_ranges = ws.merged_cells.ranges
            o = 0
            while merged_ranges:
                for entry in merged_ranges:
                    o = +1
                    print("  unMerging: " + str(o) + ": " + str(entry))
                    ws.unmerge_cells(str(entry))
        self.wb.save(self.save_name)

    def read_data(self):
        self.unmerge()
        wb = openpyxl.load_workbook(filename=self.save_name)
        return wb


class ReadWtData:
    def __init__(self):
        self.filename = "20200630.xlsx"
        self.wb = openpyxl.load_workbook(filename=self.filename)
        self.save_name = "new_wt.xlsx"

    def unmerge(self):
        for ws in self.wb:
            print(ws.title)
            merged_ranges = ws.merged_cells.ranges
            o = 0
            while merged_ranges:
                for entry in merged_ranges:
                    o = +1
                    print("  unMerging: " + str(o) + ": " + str(entry))
                    ws.unmerge_cells(str(entry))
        self.wb.save(self.save_name)

    def read_data(self):
        self.unmerge()
        wb = openpyxl.load_workbook(filename=self.save_name)
        return wb


class CalDate:
    def __init__(self):
        pass

    def calculate_date(self):
        ot_records = ReadOtData().read_data()
        # ot_name = ot_records[0]['over']
        # print(ot_name)
        ot_start = StrToDatetime(ot_records[0]['开始时间']).res
        print(ot_start)
        d_res = IsWorkDay(ot_start).res
        if d_res is True:
            print("workday")
        else:
            print("holiday")
        ot_end = StrToDatetime(ot_records[0]['结束时间']).res
        print(ot_end)
        ot_mins = ElasTime(ot_start, ot_end).res / 60
        print(ot_mins)



# CalDate().calculate_date()
#     # ReadOtData().read_xlsx()
# x = ReadWtData().read_sheet()
# print(x)


class CalWt:
    def __init__(self):
        self.wb = openpyxl.load_workbook(filename='new_wt.xlsx')

    def match_records(self, name, date):
        # print(name)
        # print(date)
        wt_records = self.wb
        for wt_ws in wt_records:
            if wt_ws.title == "每日统计":
                for rec in wt_ws.rows:
                    wt_name = rec[0].value
                    wt_date = rec[6].value
                    if (wt_name is not None) and (wt_date is not None):
                        # print("judge")
                        # print(name)
                        # print(wt_name)
                        # print(date)
                        # print(wt_date)
                        if (name in wt_name) and (date in wt_date):
                            wt_start = rec[8].value
                            wt_end = rec[10].value

                            x = ElasTime(wt_start, wt_end).elas_hrs()
                            print(x)
                        

class CalOt:
    def __init__(self):
        self.wb = openpyxl.load_workbook(filename='new_ot.xlsx')

    def calculate_date(self):
        ot_records = self.wb

        for ot_ws in ot_records:
            for rec in ot_ws.rows:
                if (rec[0].value) != "审批编号":
                    ot_name = rec[14].value
                    ot_start = rec[15].value
                    ot_end = rec[16].value
                    ot_elapsed = rec[20].value
                    ot_start_day = ot_start.split()[0][2:]
                    ot_start_time = ot_start.split()[1]
                    ot_end_day = ot_end.split()[0][2:]
                    ot_end_time = ot_end.split()[1]
                    CalWt().match_records(ot_name, ot_start_day)
                    # x = StrToDatetime(ot_start).res
                    # print(x)
                    #
                    # print(ot_name)
                    # print(ot_start)
                    # print(ot_end)
                    # print(ot_elapsed)






        # print(ot_records.worksheets[0].title)
        # ot_start = ot_records.worksheets[0]['P2'].value
        # print(ot_start)
        # ots = ot_records.worksheets[0]['P']
        # for i in ots:
        #     # if i is not ots[0]:
        #     if i.value is not '开始时间':
        #         print(i.value)
        # print(type(ot_start))
        # ot_start = StrToDatetimeHM(ot_records.worksheets[0]['p2'].value).res
        # print(ot_start)

    def unmerge(self, wb, save_name):
        for ws in wb:
            print(ws.title)
            merged_ranges = ws.merged_cells.ranges
            o = 0
            while merged_ranges:
                for entry in merged_ranges:
                    o = +1
                    print("  unMerging: " + str(o) + ": " + str(entry))
                    ws.unmerge_cells(str(entry))
        wb.save(save_name)
        
        
        
        
        
        
        
        
        


# wt_wb = ReadWtData().read_data()
# ot_wb = ReadOtData().read_data()
CalOt().calculate_date()
