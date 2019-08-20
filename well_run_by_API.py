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
def calc_well_plin_pwf_atma_for_fsolve(c_calibr_head_d):
    c_calibr_head_d = float(c_calibr_head_d)
    ESPstr = UniflocVBA.calc_ESP_encode_string(esp_id, ESP_head_nom, ESP_freq,
                                               u_motor_data_v, power_motor_nom_kwt,
                                               tsep_c, ESP_Hmes_m=h_tube_m,
                                               c_calibr_power=1,
                                               c_calibr_rate=1)
    result = UniflocVBA.calc_well_plin_pwf_atma(qliq_m3day, watercut_perc, p_wf_atm, p_cas_data_atm, Wellstr,
                                                PVTstr, ESPstr, c_calibr_head_d=c_calibr_head_d)

    plin_calculated_atm = result[0][0]
    result_for_folve = (plin_calculated_atm - p_wellhead_data_atm) ** 2
    print(plin_calculated_atm)
    return result_for_folve

fsolve(calc_well_plin_pwf_atma_for_fsolve, 0.5, xtol = 0.5)
from scipy.optimize import minimize
print("     \n")
minimize(calc_well_plin_pwf_atma_for_fsolve, [1], bounds = [[0,1]])
print(str(c_calibr_head_d) + " c_calibr_head_d")


result_fsolve = UniflocVBA.calc_well_plin_pwf_atma(qliq_m3day, watercut_perc, p_wf_atm, p_cas_data_atm, Wellstr,
                                            PVTstr, ESPstr, c_calibr_head_d=c_calibr_head_d)
print(str(p_wellhead_data_atm) +" p_wellhead_data_atm ")
for i in range(len(result_fsolve[0])):
    print(str(result_fsolve[1][i]) + " -  " + str(result_fsolve[0][i]))



