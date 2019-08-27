"""
Модуль для массового расчета скважин, оснащенных УЭЦН

Кобзарь О.С Хабибуллин Р.А. 21.08.2019
"""

import description_generated.python_api as python_api
from scipy.optimize import minimize
import pandas as pd
UniflocVBA = python_api.API("UniflocVBA_7.xlam")
import time
import sys
sys.path.append("../")
import datetime


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
        self.muob_cp = 100

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
        self.p_cas_data_atm = -1

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
        self.p_buf_data_atm = 27
        self.h_perf_m = 831
        self.h_pump_m = 830
        self.udl_m = 0

        self.ksep_d = 0.9
        self.KsepGS_fr = 0.9
        self.c_calibr_head_d = None
        self.c_calibr_rate_d = 1
        self.c_calibr_power_d = 1
        self.hydr_corr = 0
        self.result = None


def mass_calculation(well_state):
    this_state = well_state

    def calc_well_plin_pwf_atma_for_fsolve(c_calibr_head_d):
        c_calibr_power_d = c_calibr_head_d[1]
        c_calibr_head_d = c_calibr_head_d[0]
        c_calibr_rate_d = this_state.c_calibr_rate_d
               #this_state.c_calibr_power_d
        PVTstr = UniflocVBA.calc_PVT_encode_string(this_state.gamma_gas, this_state.gamma_oil,
                                                   this_state.gamma_wat, this_state.rsb_m3m3, this_state.rp_m3m3,
                                                   this_state.pb_atm, this_state.tres_c,
                                                   this_state.bob_m3m3, this_state.muob_cp,
                                                   ksep_fr=this_state.ksep_d, pksep_atma=this_state.psep_atm,
                                                   tksep_C=this_state.tsep_c)
        Wellstr = UniflocVBA.calc_well_encode_string(this_state.h_perf_m,
                                                     this_state.h_pump_m,
                                                     this_state.udl_m,
                                                     this_state.dcas_mm,
                                                     this_state.d_tube_mm,
                                                     this_state.d_choke_mm,
                                                     tbh_C=this_state.tres_c)
        ESPstr = UniflocVBA.calc_ESP_encode_string(this_state.esp_id,
                                                   this_state.ESP_head_nom,
                                                   this_state.ESP_freq,
                                                   this_state.u_motor_data_v,
                                                   this_state.power_motor_nom_kwt,
                                                   this_state.tsep_c,
                                                   t_dis_C = -1,
                                                   KsepGS_fr=this_state.KsepGS_fr,
                                                   ESP_Hmes_m=this_state.h_tube_m,
                                                   c_calibr_head=c_calibr_head_d,
                                                   c_calibr_rate=c_calibr_rate_d,
                                                   c_calibr_power=c_calibr_power_d)
        result = UniflocVBA.calc_well_plin_pwf_atma(this_state.qliq_m3day, this_state.watercut_perc,
                                                    this_state.p_wf_atm,
                                                    this_state.p_cas_data_atm, Wellstr,
                                                    PVTstr, ESPstr, this_state.hydr_corr,
                                                    this_state.ksep_d, c_calibr_head_d, c_calibr_power_d,
                                                    c_calibr_rate_d)



        this_state.result = result

        """p_buf_calc_atm = result[0][2]
        result_for_folve = (p_buf_calc_atm - this_state.p_buf_data_atm) ** 2
        print(p_buf_calc_atm)"""

        #print(this_state.result)

        p_buf_calc_atm = result[0][2]
        power_CS_calc_W = result[0][16]
        power_regulatization = 1 / 1000
        result_for_folve = (p_buf_calc_atm - this_state.p_buf_data_atm) ** 2 + \
                           (power_regulatization * (power_CS_calc_W - this_state.active_power_cs_data_kwt)) ** 2
        #print("power_CS_calc_W = " + str(power_CS_calc_W))
        #print("active_power_cs_data_kwt = " + str(this_state.active_power_cs_data_kwt))
        #print("ошибка на текущем шаге = " + str(result_for_folve))
        return result_for_folve
    result = minimize(calc_well_plin_pwf_atma_for_fsolve, [0.5, 0.5], bounds=[[0, 20], [0, 20]])

    print(result)
    #print(result.x[0])
    true_result = this_state.result
    return true_result
    #for i in range(len(true_result[0])):
    #    print(str(true_result[1][i]) + " -  " + str(true_result[0][i]))



start = datetime.datetime(2019,2,3)
end = datetime.datetime(2019,2,27)
prepared_data = pd.read_csv("stuff_to_merge/input_data.csv")
prepared_data.index = pd.to_datetime(prepared_data["Unnamed: 0"])
prepared_data = prepared_data[(prepared_data.index > start) & (prepared_data.index < end)]
del prepared_data["Unnamed: 0"]

result_list = []
result_dataframe = {'d':[2]}
result_dataframe = pd.DataFrame(result_dataframe)
start_time = time.time()
for i in range(prepared_data.shape[0]):
#for i in range(2):
    start_in_loop_time = time.time()
    row_in_prepared_data = prepared_data.iloc[i]
    this_state = all_ESP_data()
    this_state.qliq_m3day = row_in_prepared_data[' Объемный дебит жидкости']
    this_state.watercut_perc = row_in_prepared_data[' Процент обводненности']
    this_state.rp_m3m3 = row_in_prepared_data['ГФ']
    this_state.p_buf_data_atm = row_in_prepared_data['Рбуф']
    this_state.p_intake_data_atm = row_in_prepared_data[' Давление на приеме насоса (пласт. жидкость)'] * 10
    this_state.tsep_c = row_in_prepared_data[' Температура на приеме насоса (пласт. жидкость)']
    this_state.tres_c = 16
    this_state.psep_atm = row_in_prepared_data[' Давление на приеме насоса (пласт. жидкость)'] * 10
    this_state.p_wf_atm = row_in_prepared_data[' Давление на приеме насоса (пласт. жидкость)'] * 10
    this_state.active_power_cs_data_kwt = row_in_prepared_data[' Активная мощность'] * 1000
    this_result = mass_calculation(this_state)
    result_list.append(this_result)
    end_in_loop_time = time.time()
    print("Затрачено времени в итерации: " + str(i) + " - " + str(end_in_loop_time - start_in_loop_time))
    new_dict = {}
    for i in range(len(this_result[1])):
        new_dict[this_result[1][i]] = [this_result[0][i]]
        #print(str(this_result[1][i]) + " -  " + str(this_result[0][i]))
    new_dataframe = pd.DataFrame(new_dict)
    result_dataframe = result_dataframe.append(new_dataframe, sort=False)


end_time = time.time()
print("Затрачено всего: " + str(end_time - start_time))

result_dataframe.to_csv("stuff_to_merge/check_result_26_08_2019.csv")





"""def calc_well_plin_pwf_atma_for_fsolve(c_calibr_head_d):
    c_calibr_power_d = c_calibr_head_d[1]
    c_calibr_head_d = c_calibr_head_d[0]
    c_calibr_rate_d = this_state.c_calibr_rate_d
    # this_state.c_calibr_power_d
    PVTstr = UniflocVBA.calc_PVT_encode_string(this_state.gamma_gas, this_state.gamma_oil,
                                               this_state.gamma_wat, this_state.rsb_m3m3, this_state.rp_m3m3,
                                               this_state.pb_atm, this_state.tres_c,
                                               this_state.bob_m3m3, this_state.muob_cp,
                                               ksep_fr=this_state.ksep_d, pksep_atma=this_state.psep_atm,
                                               tksep_C=this_state.tsep_c)
    Wellstr = UniflocVBA.calc_well_encode_string(this_state.h_perf_m,
                                                 this_state.h_pump_m,
                                                 this_state.udl_m,
                                                 this_state.dcas_mm,
                                                 this_state.d_tube_mm,
                                                 this_state.d_choke_mm,
                                                 tbh_C=this_state.tres_c)
    ESPstr = UniflocVBA.calc_ESP_encode_string(this_state.esp_id,
                                               this_state.ESP_head_nom,
                                               this_state.ESP_freq,
                                               this_state.u_motor_data_v,
                                               this_state.power_motor_nom_kwt,
                                               this_state.tsep_c,
                                               t_dis_C=-1,
                                               KsepGS_fr=this_state.KsepGS_fr,
                                               ESP_Hmes_m=this_state.h_tube_m,
                                               c_calibr_head=c_calibr_head_d,
                                               c_calibr_rate=c_calibr_rate_d,
                                               c_calibr_power=c_calibr_power_d)
    result = UniflocVBA.calc_well_plin_pwf_atma(this_state.qliq_m3day, this_state.watercut_perc,
                                                this_state.p_wf_atm,
                                                this_state.p_cas_data_atm, Wellstr,
                                                PVTstr, ESPstr, this_state.hydr_corr,
                                                this_state.ksep_d, c_calibr_head_d, c_calibr_power_d,
                                                c_calibr_rate_d)
    return result

for i in range(1):
    start_in_loop_time = time.time()
    row_in_prepared_data = prepared_data.iloc[i]
    this_state = all_ESP_data()
    this_state.qliq_m3day = row_in_prepared_data[' Объемный дебит жидкости']
    this_state.watercut_perc = row_in_prepared_data[' Процент обводненности']
    this_state.rp_m3m3 = row_in_prepared_data['ГФ']
    this_state.p_buf_data_atm = row_in_prepared_data['Рбуф']
    this_state.p_intake_data_atm = row_in_prepared_data[' Давление на приеме насоса (пласт. жидкость)'] * 10
    this_state.tsep_c = row_in_prepared_data[' Температура на приеме насоса (пласт. жидкость)']
    this_state.tres_c = 16
    this_state.psep_atm = row_in_prepared_data[' Давление на приеме насоса (пласт. жидкость)'] * 10
    this_state.p_wf_atm = row_in_prepared_data[' Давление на приеме насоса (пласт. жидкость)'] * 10
    this_state.active_power_cs_data_kwt = row_in_prepared_data[' Активная мощность'] * 1000

    result_dataframe = {'d': list(range(50, 200, 1))}
    result_dataframe = pd.DataFrame(result_dataframe)
    for j in range(50, 200, 1):
        c_calibr_head = j /100
        one_column_data = []
        for k in range(50, 200, 1):
            c_calibr_power = k
            calibr_list = [c_calibr_head, c_calibr_power]
            result = calc_well_plin_pwf_atma_for_fsolve(calibr_list)
            p_buf_calc_atm = result[0][2]
            power_CS_calc_W = result[0][16]
            power_regulatization = 1 / 1000
            result_for_folve = (p_buf_calc_atm - this_state.p_buf_data_atm) ** 2 + \
                               (power_regulatization * (power_CS_calc_W - this_state.active_power_cs_data_kwt)) ** 2
            one_column_data.append(result_for_folve)
            #result_dataframe = result_dataframe.append(pd.DataFrame({c_calibr_head: one_column_data}), sort=False)
        result_dataframe[c_calibr_head] = one_column_data

    del result_dataframe['d']
    #result_dataframe.index = result_dataframe.columns
    #result_dataframe.to_csv(stuff_to_merge/solution_surface_2.csv) """

