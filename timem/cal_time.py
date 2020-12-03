import datetime
import chinese_calendar


class CalDatetime:
    def __init__(self, str_time):
        self.time_tr = str_time

    def dt_hrmin(self):
        hm = datetime.datetime.strptime(self.time_tr, '%H:%M')
        return hm

    def dt_ymd(self):
        ymd = datetime.datetime.strptime(self.time_tr, '%Y-%m-%d')
        return ymd

    def dt_all(self):
        all = datetime.datetime.strptime(self.time_tr, '%Y-%m-%d %H:%M')
        return all

    def dt_nextday(self):
        nextday = datetime.datetime.strptime(self.time_tr, '%H:%M') + datetime.timedelta(hours=24)
        return nextday


class ElasTime:
    def __init__(self, time_start, time_end):
        self.time_start = time_start
        self.time_end = time_end

    def hm_str_to_dt(self):
        self.time_start = CalDatetime(self.time_start).dt_hrmin()
        self.time_end = CalDatetime(self.time_end).dt_hrmin()

    def hm_str_to_next_dt(self):
        self.time_start = CalDatetime(self.time_start).dt_hrmin()
        self.time_end = CalDatetime(self.time_end).dt_nextday()

    def elas_hrs(self):
        self.hm_str_to_dt()
        es = (self.time_end - self.time_start).seconds/3600
        return es

    def next_hrs(self):
        self.hm_str_to_next_dt()
        es = (self.time_end - self.time_start).seconds / 3600
        return es


class IsWorkDay:
    def __init__(self, w_date):
        # w_date_f = datetime.datetime.strptime(w_date, '%Y-%m-%d %H:%M')

        # self.w_res = chinese_calendar.is_workday(w_date_f)
        w_date_f = CalDatetime(w_date).dt_all()
        self.w_res = chinese_calendar.is_workday(w_date_f)
        on_holiday, holiday_name = chinese_calendar.get_holiday_detail(w_date_f)
        # print("is work day?", self.w_res)
        # print("is holiday?", on_holiday)
        # print("holiday name?", holiday_name)

    @property
    def res(self):
        return self.w_res


class XTest:
    def __init__(self):
        pass

    def dtcompr(self):
        std_end = '2020-06-02 17:30'
        wt_end = '2020-06-03 17:30'
        dt_std_end = CalDatetime(std_end).dt_all()
        dt_wt_end = CalDatetime(wt_end).dt_all()
        if dt_wt_end >= dt_std_end:
            print('True')
        else:
            print('False')

    def test_str(self):
        strtest = []
        strtest.append('y')
        print('x')
        print(strtest)