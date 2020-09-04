from timem import CalDatetime, ElasTime, IsWorkDay


class OtApprv:
    def __init__(self, ot_day, ot_s, ot_e, wt_s, wt_e, wt_elas):
        self.dt_std_edn = CalDatetime('17:30').dt_hrmin()
        self.res = False
        self.ref_cmts = list()
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
            if (wt_elas >= 11.5) and (ot_s >= self.dt_std_edn) and (ot_s >= (wt_s + 9)) and (ot_e <= wt_e):
                self.res = True
            else:
                if wt_elas < 11.5:
                    self.res = False
                    self.ref_cmts.append('当日工作总时长不足')
                if ot_s < self.dt_std_edn:
                    self.res = False
                    self.ref_cmts.append('工作日加班申请开始时间早于17:30')
                if ot_s < (wt_s + 9):
                    self.res = False
                    self.ref_cmts.append('工作日加班申请开始时间早于标准下班时间')
                if ot_e > wt_e:
                    self.res = False
                    self.ref_cmts.append('加班申请结束时间' + ot_e +'晚于打卡结束时间'+ wt_e)

    @property
    def ot_res(self):
        return self.res, self.ref_cmts