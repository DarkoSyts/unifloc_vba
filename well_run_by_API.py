import description_generated.python_api as python_api
from scipy.optimize import fsolve

UniflocVBA = python_api.API("UniflocVBA_7.xlam")

gamma_oil = 0.945
gamma_gas = 0.9
gamma_wat = 1.011
rsb_m3m3 = 29.25
tres_c = 16
pb_atm = 40
bob_m3m3 = 1.045
muob_cp = 110
ksep_d = 0.7
psep_atm = 30
tsep_c = 30
d_choke_mm = 8
dcas_mm = 160
h_tube_m = 830
d_tube_mm = 75
Power_motor_nom_kWt = 140
ESP_head_nom = 1500
ESP_rate_nom = 320
ESP_freq = 38

rp_m3m3 = 30

p_intake_data_atm = 29.93
p_wellhead_data_atm = 22.70
p_buf_data_atm = 27.0
p_wf_atm = 29.93
p_cas_data_atm = 25.90

eff_motor_d = 0.89
i_motor_nom_a = 6
power_motor_nom_kwt = 140
i_motor_data_a = 42.94
cos_phi_data_d = 0.70818
load_motor_data_d = 0.55957
u_motor_data_v = 1546.65
active_power_cs_data_kwt = 81.297297


PVTstr = UniflocVBA.calc_PVT_encode_string(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3,
                                           pb_atm, tres_c, bob_m3m3, muob_cp, ksep_fr=ksep_d, pksep_atma=psep_atm,
                                           tksep_C=tsep_c)
Wellstr = UniflocVBA.calc_well_encode_string(831 ,830,0,dcas_mm,d_tube_mm,d_choke_mm, tbh_C=tres_c)

esp_id = UniflocVBA.calc_ESP_id_by_rate(320)


ESPstr = UniflocVBA.calc_ESP_encode_string(esp_id, ESP_head_nom, ESP_freq,
                                           u_motor_data_v, power_motor_nom_kwt,
                                           tsep_c, ESP_Hmes_m=h_tube_m,
                                           c_calibr_power=1,
                                           c_calibr_rate=1)
qliq_m3day = 122.2
watercut_perc = 25.6
result = UniflocVBA.calc_well_plin_pwf_atma(qliq_m3day, watercut_perc, p_wf_atm, p_cas_data_atm, Wellstr,
                                            PVTstr, ESPstr, c_calibr_head_d=1)
for i in range(len(result[0])):
    print(str(result[1][i]) + " -  " + str(result[0][i]))

c_calibr_head_d = 0.6

class all_ESP_data():
    def __init__(self):
        self.esp_id = UniflocVBA.calc_ESP_id_by_rate(320)
        self.gamma_oil = 0.945
        self.gamma_gas = 0.9
        self.gamma_wat = 1.011
        self.rsb_m3m3 = 29.25
        self.tres_c = 16
        self.pb_atm = 40
        self.bob_m3m3 = 1.045
        self.muob_cp = 400
        self.ksep_d = 0.7
        self.psep_atm = 30
        self.tsep_c = 30
        self.d_choke_mm = 8
        self.dcas_mm = 160
        self.h_tube_m = 830
        self.d_tube_mm = 75
        self.Power_motor_nom_kWt = 140
        self.ESP_head_nom = 1500
        self.ESP_rate_nom = 320
        self.ESP_freq = 38

        self.rp_m3m3 = 30

        self.p_intake_data_atm = 29.93
        self.p_wellhead_data_atm = 22.70
        self.p_buf_data_atm = 27.0
        self.p_wf_atm = 29.93
        self.p_cas_data_atm = 25.90

        self.eff_motor_d = 0.89
        self.i_motor_nom_a = 6
        self.power_motor_nom_kwt = 140
        self.i_motor_data_a = 42.94
        self.cos_phi_data_d = 0.70818
        self.load_motor_data_d = 0.55957
        self.u_motor_data_v = 1546.65
        self.active_power_cs_data_kwt = 81.297297

        self.qliq_m3day = 122.2
        self.watercut_perc = 25.6



this_state = all_ESP_data()


def calc_well_plin_pwf_atma_for_fsolve(c_calibr_head_d):
    c_calibr_head_d = float(c_calibr_head_d)
    Wellstr = UniflocVBA.calc_well_encode_string(831, 830, 0, this_state.dcas_mm, this_state.d_tube_mm,
                                                 this_state.d_choke_mm, tbh_C=this_state.tres_c)
    ESPstr = UniflocVBA.calc_ESP_encode_string(this_state.esp_id, this_state.ESP_head_nom, this_state.ESP_freq,
                                               this_state.u_motor_data_v, this_state.power_motor_nom_kwt,
                                               this_state.tsep_c, ESP_Hmes_m=this_state.h_tube_m,
                                               c_calibr_power=1,
                                               c_calibr_rate=1)
    result = UniflocVBA.calc_well_plin_pwf_atma(this_state.qliq_m3day, this_state.watercut_perc, this_state.p_wf_atm,
                                                this_state.p_cas_data_atm, Wellstr,
                                                PVTstr, ESPstr, c_calibr_head_d=c_calibr_head_d)

    p_buf_calc_atm = result[0][2]
    result_for_folve = (p_buf_calc_atm - p_buf_data_atm) ** 2
    print(p_buf_calc_atm)
    print(result_for_folve)
    return result_for_folve

#fsolve(calc_well_plin_pwf_atma_for_fsolve, 0.5, xtol = 0.5)
from scipy.optimize import minimize
print("     \n")
result = minimize(calc_well_plin_pwf_atma_for_fsolve, [0.5], bounds = [[0,1]])
print(result)
print(result.x[0])
print(str(c_calibr_head_d) + " c_calibr_head_d")


result_fsolve = UniflocVBA.calc_well_plin_pwf_atma(qliq_m3day, watercut_perc, p_wf_atm, p_cas_data_atm, Wellstr,
                                            PVTstr, ESPstr, c_calibr_head_d=result.x[0])
print(str(p_wellhead_data_atm) +" p_wellhead_data_atm ")
for i in range(len(result_fsolve[0])):
    print(str(result_fsolve[1][i]) + " -  " + str(result_fsolve[0][i]))

'''import pandas as pd

class well():
    def __init__(self):
        self.result = None
        self.gamma_oil = 0.945
        self.gamma_gas = 0.9
        self.gamma_wat = 1.011
        self.rsb_m3m3 = 29.25
        self.tres_c = 16
        self.pb_atm = 40
        self.bob_m3m3 = 1.045
        self.muob_cp = 400
        self.ksep_d = 0.7
        self.psep_atm = 30
        self.tsep_c = 30
        self.d_choke_mm = 8
        self.dcas_mm = 160
        self.h_tube_m = 830
        self.d_tube_mm = 75
        self.Power_motor_nom_kWt = 140
        self.ESP_head_nom = 1500
        self.ESP_rate_nom = 320
        self.ESP_freq = 38

        self.rp_m3m3 = 30

        self.p_intake_data_atm = 29.93
        self.p_wellhead_data_atm = 22.70
        self.p_buf_data_atm = 27.0
        self.p_wf_atm = 29.93
        self.p_cas_data_atm = 25.90

        self.eff_motor_d = 0.89
        self.i_motor_nom_a = 6
        self.power_motor_nom_kwt = 140
        self.i_motor_data_a = 42.94
        self.cos_phi_data_d = 0.70818
        self.load_motor_data_d = 0.55957
        self.u_motor_data_v = 1546.65
        self.active_power_cs_data_kwt = 81.297297

        self.qliq_m3day = 122.2
        self.watercut_perc = 25.6

        self.c_calibr_head_d = 1

        self.PVTstr = None
        self.Wellstr = None
        self.Wellstr = None
        self.esp_id = None
        self.ESPstr = None
        self.p_buf_calculated_atm = None

    def calc_well_plin_pwf_atma_for_fsolve(self, c_calibr_head_d):
        c_calibr_head_d = float(c_calibr_head_d)
        self.PVTstr = UniflocVBA.calc_PVT_encode_string(self.gamma_gas, self.gamma_oil, self.gamma_wat, self.rsb_m3m3,
                                                   self.rp_m3m3,
                                                   self.pb_atm, self.tres_c, self.bob_m3m3, self.muob_cp,
                                                   ksep_fr=self.ksep_d,
                                                   pksep_atma=self.psep_atm,
                                                   tksep_C=self.tsep_c)
        self.Wellstr = UniflocVBA.calc_well_encode_string(831, 830, 0, self.dcas_mm, self.d_tube_mm, self.d_choke_mm,
                                                     tbh_C=self.tres_c)

        self.esp_id = UniflocVBA.calc_ESP_id_by_rate(320)

        self.ESPstr = UniflocVBA.calc_ESP_encode_string(self.esp_id, self.ESP_head_nom, self.ESP_freq,
                                                   self.u_motor_data_v, self.power_motor_nom_kwt,
                                                   self.tsep_c, ESP_Hmes_m=self.h_tube_m,
                                                   c_calibr_power=1,
                                                   c_calibr_rate=1)
        self.result = UniflocVBA.calc_well_plin_pwf_atma(self.qliq_m3day, self.watercut_perc, self.p_wf_atm,
                                                         self.p_cas_data_atm, self.Wellstr,
                                                    self.PVTstr, self.ESPstr, c_calibr_head_d=c_calibr_head_d)

        self.p_buf_calculated_atm = result[0][3]
        result_out = (self.p_buf_calculated_atm - self.p_buf_data_atm) ** 2
        print(result_out)
        return result_out

    def calc_well(self):
        #fsolve(self.calc_well_plin_pwf_atma_for_fsolve, 0.6, xtol=1)
        minimize(self.calc_well_plin_pwf_atma_for_fsolve, [0.6], bounds=[[0,100]])

well_1354 = well()
well_1354.calc_well()
print(well_1354.result)

check = UniflocVBA.book.macro('check_minimaze')
print(check(2))
def check_for_minimaze(x):
    x = float(x)
    result = check(x)
    return result
class check_minimaze():
    def __init__(self):
        self.func_result = None
        self.x = None
        self.check = UniflocVBA.book.macro('check_minimaze')
    def calc(self, x):
        self.x = float(x)
        print('self.x')
        self.func_result = self.check(x)
        print(self.func_result)
        return self.func_result
result = minimize(check_for_minimaze, [2], bounds=[[0,100]])
print(result.x)

check2 = check_minimaze()
minimize(check2.calc, [2], bounds=[[0,100]])
print(result.x) '''