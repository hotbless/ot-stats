from timem import CalDatetime, ElasTime, IsWorkDay


class OtApprv:
    def __init__(self, ot_day, ot_s, ot_e, wt_s, wt_e, wt_elas):
        self.dt_std_edn = CalDatetime('17:30').dt_hrmin()
        self.res = False
        ot_elas_hrs = ElasTime(ot_s, ot_e).elas_hrs()
        ot_is_wd = IsWorkDay(ot_day).res
        print(ot_day + ' 是工作日: ' + str(ot_is_wd))
        if ot_is_wd is False:
            if (ot_elas_hrs >= 2.5) and (ot_s >= wt_s) and (ot_e >= wt_e):
                self.res = True
        else:
            if (wt_elas >= 11.5) and (ot_s >= self.dt_std_edn) and (ot_s >= (wt_s + 9)) and (ot_e >= wt_e):
                self.res = True

    @property
    def ot_res(self):
        return self.res