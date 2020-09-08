from timem import CalDatetime, ElasTime, IsWorkDay
import datetime


class OtApprv:
    def __init__(self, ot_day, ot_s, ot_e, wt_s, wt_e, wt_elas):
        std_et = '17:30'
        self.dt_std_edn = CalDatetime(std_et).dt_hrmin()
        self.res = False
        self.ref_cmts = list()
        ot_s_dt = CalDatetime(ot_s).dt_hrmin()
        wt_s_dt = CalDatetime(wt_s).dt_hrmin()
        ot_elas_hrs = ElasTime(ot_s, ot_e).elas_hrs()
        ot_is_wd = IsWorkDay(ot_day).res
        print(ot_day + ' 是工作日: ' + str(ot_is_wd))
        if ot_is_wd is False:
            if (ot_elas_hrs >= 2.5) and (ot_s >= wt_s) and (ot_e <= wt_e):
                self.res = True
            else:
                if ot_elas_hrs < 2.5:
                    self.res = False
                    self.ref_cmts.append('加班时长小于2.5小时')
                if ot_s < wt_s:
                    self.res = False
                    self.ref_cmts.append('节假日加班申请开始时间早于打卡开始时间')
                if ot_e > wt_e:
                    self.res = False
                    self.ref_cmts.append('节假日加班申请结束时间' + ot_e +'晚于打卡结束时间'+ wt_e)
        else:
            if (wt_elas >= 11.5) and (ot_s >= std_et) and (ot_s_dt >= (wt_s_dt + datetime.timedelta(hours=9))) and (ot_e <= wt_e):
                self.res = True
            else:
                if wt_elas < 11.5:
                    self.res = False
                    self.ref_cmts.append('当日工作总时长' + str(wt_elas) + '不足11.5')
                if ot_s < std_et:
                    self.res = False
                    self.ref_cmts.append('工作日加班申请开始时间' + ot_s + '早于17:30')
                if ot_s_dt < (wt_s_dt + datetime.timedelta(hours=9)):
                    cal_wt_e = (wt_s_dt + datetime.timedelta(hours=9)).strftime('%H:%M')
                    self.res = False
                    self.ref_cmts.append('工作日加班申请开始时间' + ot_s + '早于所应下班时间' + cal_wt_e)
                if ot_e > wt_e:
                    self.res = False
                    self.ref_cmts.append('加班申请结束时间' + ot_e + '晚于打卡结束时间' + wt_e)

    @property
    def ot_res(self):
        return self.res, self.ref_cmts