import datetime
from timem import CalDatetime, ElasTime, IsWorkDay

class TestTime:
    def __init__(self):
        pass

    def test(self):
        ot_s = '18:00'
        ot_e = '20:30'
        wt_s = '08:54'
        wt_e = '20:30'

        dt_std_edn = CalDatetime('17:30').dt_hrmin()
        wt_e_std = CalDatetime(wt_e).dt_hrmin()
        wt_s_std = CalDatetime(wt_s).dt_hrmin()
        print(dt_std_edn)
        print(wt_e_std)
        print(wt_s_std)


        print('ot_s > ot_e' + ':'+str(ot_s > ot_e))
        print('wt_e >= ot_e' + ':' + str(wt_e >= ot_e))
        print('wt_s < ot_s' + ':' + str(wt_s < ot_s))
        print('wt_std < dt_std' + ':' + str(wt_e_std < dt_std_edn))
        print((wt_s_std + datetime.timedelta(hours=9)).strftime('%H:%M'))



TestTime().test()