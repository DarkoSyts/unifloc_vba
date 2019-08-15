import unittest
from description_generated import python_api

UniflocVBA = python_api.API("UniflocVBA_7.xlam")
UniflocVBA.calc_PVT_bg_m3m3(30,30)
p_atm = 30
t_c = 30
delta_1_in_test = 0.00001
q_liq_m3day = 100
fw_perc = 20
d_choke_mm = 10
p_in_atma = 60
p_out_atma = 30
free_gas_d = 0.3
ESP_string = UniflocVBA.calc_ESP_encode_string(esp_ID = 747)
PVT_string = UniflocVBA.calc_PVT_encode_string()
Well_string = UniflocVBA.calc_well_encode_string()
qgas_sm3day = 100 * 100

class TestPVT(unittest.TestCase):
    def test_PVT_bg_m3m3(self):
        result = UniflocVBA.calc_PVT_bg_m3m3(p_atm, t_c)
        self.assertAlmostEqual(result, 0.03333992560249548, delta=delta_1_in_test)

    def test_PVT_bo_m3m3(self):
        result = UniflocVBA.calc_PVT_bo_m3m3(p_atm, t_c)
        self.assertAlmostEqual(result, 1.031670546292832, delta=delta_1_in_test)

    def test_PVT_bw_m3m3(self):
        result = UniflocVBA.calc_PVT_bw_m3m3(p_atm, t_c)
        self.assertAlmostEqual(result, 1.00512664544482, delta=delta_1_in_test)

    def test_PVT_mu_gas_cP(self):
        result = UniflocVBA.calc_PVT_mu_gas_cP(p_atm, t_c)
        self.assertAlmostEqual(result, 0.011841488533836212, delta=delta_1_in_test)

    def test_PVT_mu_oil_cP(self):
        result = UniflocVBA.calc_PVT_mu_oil_cP(p_atm, t_c)
        self.assertAlmostEqual(result, 10.2988304561591, delta=delta_1_in_test)

    def test_PVT_mu_wat_cP(self):
        result = UniflocVBA.calc_PVT_mu_wat_cP(p_atm, t_c)
        self.assertAlmostEqual(result, 0.7643648576261103, delta=delta_1_in_test)

    def test_PVT_rhog_kgm3(self):
        result = UniflocVBA.calc_PVT_rhog_kgm3(p_atm, t_c)
        self.assertAlmostEqual(result, 21.986251821303817, delta=delta_1_in_test)

    def test_PVT_rhoo_kgm3(self):
        result = UniflocVBA.calc_PVT_rhoo_kgm3(p_atm, t_c)
        self.assertAlmostEqual(result, 842.5664535063798, delta=delta_1_in_test)

    def test_PVT_rhow_kgm3(self):
        result = UniflocVBA.calc_PVT_rhow_kgm3(p_atm, t_c)
        self.assertAlmostEqual(result, 994.8995029949174, delta=delta_1_in_test)

    def test_PVT_pb_atma(self):
        result = UniflocVBA.calc_PVT_pb_atma(t_c)
        self.assertAlmostEqual(result, 166.54738664310437, delta=delta_1_in_test)

    def test_PVT_rs_m3m3(self):
        result = UniflocVBA.calc_PVT_rs_m3m3(p_atm, t_c)
        self.assertAlmostEqual(result, 12.620383314153672, delta=delta_1_in_test)

    def test_PVT_salinity_ppm(self):
        result = UniflocVBA.calc_PVT_salinity_ppm(p_atm, t_c)
        self.assertAlmostEqual(result, 1363.1482481105195, delta=delta_1_in_test)

    def test_PVT_STliqgas_Nm(self):
        result = UniflocVBA.calc_PVT_STliqgas_Nm(p_atm, t_c)
        self.assertAlmostEqual(result, 0.0421964516056052, delta=delta_1_in_test)

    def test_PVT_SToilgas_Nm(self):
        result = UniflocVBA.calc_PVT_SToilgas_Nm(p_atm, t_c)
        self.assertAlmostEqual(result, 0.0421964516056052, delta=delta_1_in_test)

    def test_PVT_STwatgas_Nm(self):
        result = UniflocVBA.calc_PVT_STwatgas_Nm(p_atm, t_c)
        self.assertAlmostEqual(result, 0.06025595860743578, delta=delta_1_in_test)

    def test_PVT_z(self):
        result = UniflocVBA.calc_PVT_z(p_atm, t_c)
        self.assertAlmostEqual(result, 0.9632857422186332, delta=delta_1_in_test)

    def test_PVT_decode_string(self): # TODO исправить
        result = UniflocVBA.calc_PVT_decode_string()
        self.assertAlmostEqual(result, 12.620383314153672, delta=delta_1_in_test)

    def test_PVT_encode_string(self):  # TODO исправить
        result = UniflocVBA.calc_ESP_encode_string()
        self.assertAlmostEqual(result, 12.620383314153672, delta=delta_1_in_test)

class TestMF(unittest.TestCase):
    def test_MF_gas_fraction_d(self):
        result = UniflocVBA.calc_MF_gas_fraction_d(p_atm, t_c)
        self.assertAlmostEqual(result, 0.6132212223600052, delta=delta_1_in_test)

    def test_MF_CJT_Katm(self):
        result = UniflocVBA.calc_MF_CJT_Katm(p_atm, t_c)
        self.assertAlmostEqual(result, 0.03498879664846513, delta=delta_1_in_test)

    def test_MF_dp_choke_atm(self): # TODO исправить
        result = sum(UniflocVBA.calc_MF_dp_choke_atm(q_liq_m3day, fw_perc, d_choke_mm)[0])
        self.assertAlmostEqual(result, 12.620383314153672, delta=delta_1_in_test)

    def test_MF_p_choke_atma(self):
        result = sum(UniflocVBA.calc_MF_p_choke_atma(q_liq_m3day, fw_perc, d_choke_mm)[0])
        self.assertAlmostEqual(result, 0.03498879664846513, delta=delta_1_in_test)

    def test_MF_qliq_choke_sm3day(self):
        result = sum(UniflocVBA.calc_MF_qliq_choke_sm3day(fw_perc, d_choke_mm, p_in_atma, p_out_atma)[0])
        self.assertAlmostEqual(result, 583.7511058679155, delta=delta_1_in_test)

    def test_MF_mu_mix_cP(self):
        result = UniflocVBA.calc_MF_mu_mix_cP(q_liq_m3day, fw_perc, p_atm, t_c)
        self.assertAlmostEqual(result, 0.03498879664846513, delta=delta_1_in_test)

    def test_MF_rhomix_kgm3(self):
        result = UniflocVBA.calc_MF_rhomix_kgm3(q_liq_m3day, fw_perc, p_atm, t_c)
        self.assertAlmostEqual(result, 0.03498879664846513, delta=delta_1_in_test)

    def test_MF_q_mix_rc_m3day(self):
        result = UniflocVBA.calc_MF_q_mix_rc_m3day(q_liq_m3day, fw_perc, p_atm, t_c)
        self.assertAlmostEqual(result, 0.03498879664846513, delta=delta_1_in_test)

    def test_MF_p_gas_fraction_atma(self):
        result = UniflocVBA.calc_MF_p_gas_fraction_atma(free_gas_d,  t_c, fw_perc)
        self.assertAlmostEqual(result, 44.34814453125, delta=delta_1_in_test)

    def test_MF_rp_gas_fraction_m3m3(self):
        result = UniflocVBA.calc_MF_rp_gas_fraction_m3m3(free_gas_d,  p_atm, t_c, fw_perc)
        self.assertAlmostEqual(result, 49.346923828125, delta=delta_1_in_test)

    def test_MF_gasseparator_name(self):
        result = UniflocVBA.calc_MF_gasseparator_name(1)
        self.assertAlmostEqual(result, 0.03498879664846513, delta=delta_1_in_test)

    def test_MF_ksep_gasseparator_d(self):
        result = UniflocVBA.calc_MF_ksep_gasseparator_d(1, free_gas_d, q_liq_m3day, 20)
        self.assertAlmostEqual(result, 0.789426274676, delta=delta_1_in_test)

    def test_MF_ksep_natural_d(self):
        result = UniflocVBA.calc_MF_ksep_natural_d(q_liq_m3day, fw_perc, p_atm, t_c)
        self.assertAlmostEqual(result, 0.5407649756774457, delta=delta_1_in_test)

    def test_MF_ksep_total_d(self):
        result = UniflocVBA.calc_MF_ksep_total_d(0.5, 0.9)
        self.assertAlmostEqual(result, 0.95, delta=delta_1_in_test)

    def test_MF_dp_pipe_atm(self):
        result = sum(UniflocVBA.calc_MF_dp_pipe_atm(q_liq_m3day, fw_perc, 100, p_in_atma, True))
        self.assertAlmostEqual(result, 0.6411470319037846, delta=delta_1_in_test)

    def test_MF_p_pipe_atma(self):
        result = sum(UniflocVBA.calc_MF_p_pipe_atma(q_liq_m3day, fw_perc, 100, p_in_atma, True))
        self.assertAlmostEqual(result,109.35885296809622, delta=delta_1_in_test)

    def test_MF_dpdl_atmm(self):
        result = sum(UniflocVBA.calc_MF_dpdl_atmm(70,p_atm, q_liq_m3day, q_liq_m3day * 10))
        self.assertAlmostEqual(result,103.16432810349251, delta=delta_1_in_test)

    def test_MF_p_pipe_znlf_atma(self):
        result = sum(UniflocVBA.calc_MF_p_pipe_znlf_atma(q_liq_m3day, fw_perc, 100, p_in_atma, True))
        self.assertAlmostEqual(result,104.1089549301434, delta=delta_1_in_test)

    def test_MF_calibr_pipe_m3day(self):
        result = sum(UniflocVBA.calc_MF_calibr_pipe_m3day(q_liq_m3day, fw_perc, 100, p_in_atma, p_out_atma)[0])
        self.assertAlmostEqual(result,109.35885296809622, delta=delta_1_in_test)

    def test_MF_calibr_choke_fr(self):
        result = sum(UniflocVBA.calc_MF_calibr_choke_fr(q_liq_m3day, fw_perc, d_choke_mm, p_in_atma, p_out_atma)[0])
        self.assertAlmostEqual(result, 110.4230555941965, delta=delta_1_in_test)

class TestESP(unittest.TestCase):
    def test_ESP_encode_string(self):
        result = UniflocVBA.calc_ESP_encode_string()
        self.assertAlmostEqual(result, 0.6132212223600052, delta=delta_1_in_test)

    def test_ESP_calibr_calc(self):
        result = UniflocVBA.calc_ESP_calibr_calc(q_liq_m3day, fw_perc, 30, 100, PVT_string, ESP_string)
        self.assertAlmostEqual(result, 0.6132212223600052, delta=delta_1_in_test)

    def test_ESP_decode_string(self):
        result = UniflocVBA.calc_ESP_decode_string(ESP_string)
        self.assertAlmostEqual(result, 0.6132212223600052, delta=delta_1_in_test)

    def test_ESP_dp_atm(self):
        result = sum(UniflocVBA.calc_ESP_dp_atm(q_liq_m3day, fw_perc, 30, pump_id=747))
        self.assertAlmostEqual(result, 137.991989734772, delta=delta_1_in_test)

    def test_ESP_eff_fr(self):
        result = UniflocVBA.calc_ESP_eff_fr(150, pump_id=747) # TODO не работает с насосом по умолчанию
        self.assertAlmostEqual(result, 0.5530630058389673, delta=delta_1_in_test)

    def test_ESP_head_m(self):
        result = UniflocVBA.calc_ESP_head_m(150, pump_id=747)
        self.assertAlmostEqual(result, 7.39794452714204, delta=delta_1_in_test)

    def test_ESP_id_by_rate(self):
        result = UniflocVBA.calc_ESP_id_by_rate(100)
        self.assertAlmostEqual(result, 737, delta=delta_1_in_test)

    def test_ESP_max_rate_m3day(self):
        result = UniflocVBA.calc_ESP_max_rate_m3day(pump_id=747)
        self.assertAlmostEqual(result, 355.0, delta=delta_1_in_test)

    def test_ESP_name(self):
        result = UniflocVBA.calc_ESP_name()
        self.assertAlmostEqual(result, 0.6132212223600052, delta=delta_1_in_test)

    def test_ESP_optRate_m3day(self):
        result = UniflocVBA.calc_ESP_optRate_m3day(pump_id=747)
        self.assertAlmostEqual(result, 159.0, delta=delta_1_in_test)

    def test_ESP_p_atm(self):
        result = sum(UniflocVBA.calc_ESP_p_atma(q_liq_m3day, 20, 300, pump_id=747))
        self.assertAlmostEqual(result, 584.5820558406429, delta=delta_1_in_test)

    def test_ESP_power_W(self):
        result = UniflocVBA.calc_ESP_power_W(q_liq_m3day, pump_id=747)
        self.assertAlmostEqual(result, 197.11598932884485, delta=delta_1_in_test)

    def test_ESP_system_calc(self):
        result = UniflocVBA.calc_ESP_system_calc(q_liq_m3day, fw_perc, 30,PVT_string, ESP_string)
        self.assertAlmostEqual(result, 0.6132212223600052, delta=delta_1_in_test)

class TestReservoir(unittest.TestCase):
    def test_IPR_pi_sm3dayatm(self):
        result = UniflocVBA.calc_IPR_pi_sm3dayatm(30, 100,250)
        self.assertAlmostEqual(result, 0.2405968502486054, delta=delta_1_in_test)

    def test_IPR_pwf_atma(self):
        result = UniflocVBA.calc_IPR_pwf_atma(1, 250, 100, 20, 10)
        self.assertAlmostEqual(result, 150.0, delta=delta_1_in_test)

    def test_IPR_qliq_sm3day(self):
        result = UniflocVBA.calc_IPR_qliq_sm3day(1, 250, 100)
        self.assertAlmostEqual(result, 124.68991164681252, delta=delta_1_in_test)

class TestWell(unittest.TestCase):
    def test_well_pintake_pwf_atma(self):
        result = UniflocVBA.calc_well_pintake_pwf_atma(q_liq_m3day, fw_perc, 100, Well_string, PVT_string )
        self.assertAlmostEqual(result, 84.22921891416985, delta=delta_1_in_test)

    def test_well_plin_pwf_atma(self):
        result = sum(UniflocVBA.calc_well_plin_pwf_atma(q_liq_m3day, fw_perc, 200, str_ESP= ESP_string)[0][0:-1])
        self.assertAlmostEqual(result, 2331.1486659864518, delta=delta_1_in_test)

    def test_well_pwf_Hdyn_atma(self):
        result = UniflocVBA.calc_well_pwf_Hdyn_atma(q_liq_m3day, fw_perc, 20, 700)[0][3]
        self.assertAlmostEqual(result, 19.6347298261911, delta=delta_1_in_test)

    def test_well_pwf_plin_atma(self):
        result = sum(UniflocVBA.calc_well_pwf_plin_atma(q_liq_m3day, fw_perc, 20 )[0][0:16])
        self.assertAlmostEqual(result, 539.7268628016981, delta=delta_1_in_test)

    def test_well_encode_string(self):
        result = UniflocVBA.calc_well_encode_string()
        self.assertAlmostEqual(result, 124.68991164681252, delta=delta_1_in_test)

    def test_well_decode_string(self):
        result = UniflocVBA.calc_well_decode_string(Well_string)
        self.assertAlmostEqual(result, 124.68991164681252, delta=delta_1_in_test)

class TestGLV(unittest.TestCase):  # TODO доделать последние функции для газлифта, некоторые вообще не считают
    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    """def test_GLV_IPO_p_atma(self): # не считает
        result = UniflocVBA.calc_GLV_IPO_p_atma(100, 3, 110, qgas_sm3day, 20)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)

    def test_GLV_d_choke_mm(self):
        result = UniflocVBA.calc_GLV_d_choke_mm(qgas_sm3day, 100, 80)
        self.assertAlmostEqual(result, 2.9374985373041227, delta=delta_1_in_test)"""








