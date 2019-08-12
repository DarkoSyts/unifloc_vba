H_CORRELATION = 0 # 0 - BeggsBrill, 1 - Ansari and so on 
PVT_CORRELATION = 0 # 0 -Standing, 1 -McCain, 2 - linear 
PVT_DEFAULT = "gamma_gas:0,900;gamma_oil:0,750;gamma_wat:1,000;rsb_m3m3:100,000;rp_m3m3:-1,000;pb_atma:-1,000;tres_C:90,000;bob_m3m3:-1,000;muob_cP:-1,000;PVTcorr:0;ksep_fr:0,000;pksep_atma:-1,000;tksep_C:-1,000; " 
ESP_DEFAULT = "ESP_ID:1006.00000;HeadNom_m:2000.00000;ESPfreq_Hz:50.00000;ESP_U_V:1000.00000;MotorPowerNom_kW:30.00000;Tintake_C:85.00000;t_dis_C:25.00000;KsepGS_fr:0.00000;ESP_energy_fact_Whday:0.00000;ESP_cable_type:0;ESP_Hmes_m:0.00000;ESP_gas_degradation_type:0;c_calibr_head:0.00000;PKV_work_min:-1,00000;PKV_stop_min:-1,00000;"
WELL_DEFAULT = "hperf_m:2000,00000;hpump_m:1800,00000;udl_m:0,00000;d_cas_mm:150,00000;dtub_mm:72,00000;dchoke_mm:15,00000;roughness_m:0,00010;tbh_C:85,00000;twh_C:25,00000;"
WELL_GL_DEFAULT = "hperf_m:2500,00000;htub_m:2000,00000;udl_m:0,00000;d_cas_mm:125,00000;dtub_mm:62,00000;dchoke_mm:15,00000;roughness_m:0,00010;tbh_C:100,00000;twh_C:50,00000;GLV:1;H_glv_m:1500,000;d_glv_mm:5,000;p_glv_atma:50,000;"
const_gg_ = 0.6 
const_gw_ = 1 
const_go_ = 0.86 
const_sigma_wat_gas_Nm = 0.01 
const_sigma_oil_Nm = 0.025 
const_mu_w = 0.36
const_mu_g = 0.0122 
const_mu_o = 0.7 
const_rsb_default = 100 
const_Bob_default = 1.2 
const_tres_default = 90 
const_Roughness_default = 0.0001 
Standing_based = 0 
import xlwings as xw
addin_name_str = "UniflocVBA_7.xlam"
class API():
    def __init__(self, addin_name_str):
        self.book = xw.Book(addin_name_str)
    def calc_PVT_bg_m3m3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_bg_m3m3 = self.book.macro("PVT_bg_m3m3")
        return self.PVT_bg_m3m3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_bo_m3m3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_bo_m3m3 = self.book.macro("PVT_bo_m3m3")
        return self.PVT_bo_m3m3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_bw_m3m3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_bw_m3m3 = self.book.macro("PVT_bw_m3m3")
        return self.PVT_bw_m3m3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_salinity_ppm(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_salinity_ppm = self.book.macro("PVT_salinity_ppm")
        return self.PVT_salinity_ppm(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_mu_oil_cP(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_mu_oil_cP = self.book.macro("PVT_mu_oil_cP")
        return self.PVT_mu_oil_cP(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_mu_gas_cP(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_mu_gas_cP = self.book.macro("PVT_mu_gas_cP")
        return self.PVT_mu_gas_cP(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_mu_wat_cP(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_mu_wat_cP = self.book.macro("PVT_mu_wat_cP")
        return self.PVT_mu_wat_cP(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_rs_m3m3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_rs_m3m3 = self.book.macro("PVT_rs_m3m3")
        return self.PVT_rs_m3m3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_z(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_z = self.book.macro("PVT_z")
        return self.PVT_z(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_rhoo_kgm3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_rhoo_kgm3 = self.book.macro("PVT_rhoo_kgm3")
        return self.PVT_rhoo_kgm3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_rhog_kgm3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_rhog_kgm3 = self.book.macro("PVT_rhog_kgm3")
        return self.PVT_rhog_kgm3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_rhow_kgm3(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_rhow_kgm3 = self.book.macro("PVT_rhow_kgm3")
        return self.PVT_rhow_kgm3(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_pb_atma(self, t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_pb_atma = self.book.macro("PVT_pb_atma")
        return self.PVT_pb_atma(t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_SToilgas_Nm(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_SToilgas_Nm = self.book.macro("PVT_SToilgas_Nm")
        return self.PVT_SToilgas_Nm(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_STwatgas_Nm(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_STwatgas_Nm = self.book.macro("PVT_STwatgas_Nm")
        return self.PVT_STwatgas_Nm(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_PVT_STliqgas_Nm(self, p_atma,t_C,gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,str_PVT=""):
        self.PVT_STliqgas_Nm = self.book.macro("PVT_STliqgas_Nm")
        return self.PVT_STliqgas_Nm(p_atma,t_C,gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,str_PVT)

    def calc_MF_CJT_Katm(self, p_atma,t_C,str_PVT=PVT_DEFAULT,qliq_sm3day=10,fw_perc=0):
        self.MF_CJT_Katm = self.book.macro("MF_CJT_Katm")
        return self.MF_CJT_Katm(p_atma,t_C,str_PVT,qliq_sm3day,fw_perc)

    def calc_MF_q_mix_rc_m3day(self, qliq_sm3day,fw_perc,p_atma,t_C,str_PVT=""):
        self.MF_q_mix_rc_m3day = self.book.macro("MF_q_mix_rc_m3day")
        return self.MF_q_mix_rc_m3day(qliq_sm3day,fw_perc,p_atma,t_C,str_PVT)

    def calc_MF_rhomix_kgm3(self, qliq_sm3day,fw_perc,p_atma,t_C,str_PVT=""):
        self.MF_rhomix_kgm3 = self.book.macro("MF_rhomix_kgm3")
        return self.MF_rhomix_kgm3(qliq_sm3day,fw_perc,p_atma,t_C,str_PVT)

    def calc_MF_mu_mix_cP(self, qliq_sm3day,fw_perc,p_atma,t_C,str_PVT=""):
        self.MF_mu_mix_cP = self.book.macro("MF_mu_mix_cP")
        return self.MF_mu_mix_cP(qliq_sm3day,fw_perc,p_atma,t_C,str_PVT)

    def calc_MF_gas_fraction_d(self, p_atma,t_C,fw_perc=0,str_PVT=PVT_DEFAULT):
        self.MF_gas_fraction_d = self.book.macro("MF_gas_fraction_d")
        return self.MF_gas_fraction_d(p_atma,t_C,fw_perc,str_PVT)

    def calc_MF_p_gas_fraction_atma(self, FreeGas_d,t_C,fw_perc,str_PVT=PVT_DEFAULT):
        self.MF_p_gas_fraction_atma = self.book.macro("MF_p_gas_fraction_atma")
        return self.MF_p_gas_fraction_atma(FreeGas_d,t_C,fw_perc,str_PVT)

    def calc_MF_rp_gas_fraction_m3m3(self, FreeGas_d,p_atma,t_C,fw_perc,str_PVT=PVT_DEFAULT):
        self.MF_rp_gas_fraction_m3m3 = self.book.macro("MF_rp_gas_fraction_m3m3")
        return self.MF_rp_gas_fraction_m3m3(FreeGas_d,p_atma,t_C,fw_perc,str_PVT)

    def calc_MF_ksep_natural_d(self, qliq_sm3day,fw_perc,p_intake_atma,t_intake_C=50,d_intake_mm=90,d_cas_mm=120,str_PVT=PVT_DEFAULT):
        self.MF_ksep_natural_d = self.book.macro("MF_ksep_natural_d")
        return self.MF_ksep_natural_d(qliq_sm3day,fw_perc,p_intake_atma,t_intake_C,d_intake_mm,d_cas_mm,str_PVT)

    def calc_MF_ksep_total_d(self, SepNat,SepGasSep):
        self.MF_ksep_total_d = self.book.macro("MF_ksep_total_d")
        return self.MF_ksep_total_d(SepNat,SepGasSep)

    def calc_MF_ksep_gasseparator_d(self, gsep_type_TYPE,gas_frac_d,qliq_sm3day,freq_Hz=50):
        self.MF_ksep_gasseparator_d = self.book.macro("MF_ksep_gasseparator_d")
        return self.MF_ksep_gasseparator_d(gsep_type_TYPE,gas_frac_d,qliq_sm3day,freq_Hz)

    def calc_MF_gasseparator_name(self, gsep_type_TYPE):
        self.MF_gasseparator_name = self.book.macro("MF_gasseparator_name")
        return self.MF_gasseparator_name(gsep_type_TYPE)

    def calc_MF_dpdl_atmm(self, d_m,p_atma,Ql_rc_m3day,Qg_rc_m3day,mu_oil_cP=const_mu_o,mu_gas_cP=const_mu_g,sigma_oil_gas_Nm=const_sigma_oil_Nm,gamma_oil=const_go_,gamma_gas=const_gg_,eps_m=0.0001,theta_deg=90,ZNLF=False):
        self.MF_dpdl_atmm = self.book.macro("MF_dpdl_atmm")
        return self.MF_dpdl_atmm(d_m,p_atma,Ql_rc_m3day,Qg_rc_m3day,mu_oil_cP,mu_gas_cP,sigma_oil_gas_Nm,gamma_oil,gamma_gas,eps_m,theta_deg,ZNLF)

    def calc_MF_dp_pipe_atm(self, qliq_sm3day,fw_perc,length_m,pcalc_atma,calc_along_flow,str_PVT=PVT_DEFAULT,theta_deg=90,d_mm=60,hydr_corr=H_CORRELATION,tcalc_C=50,tother_C=-1,c_calibr_grav=1,c_calibr_fric=1,roughness_m=0.0001,q_gas_sm3day=0):
        self.MF_dp_pipe_atm = self.book.macro("MF_dp_pipe_atm")
        return self.MF_dp_pipe_atm(qliq_sm3day,fw_perc,length_m,pcalc_atma,calc_along_flow,str_PVT,theta_deg,d_mm,hydr_corr,tcalc_C,tother_C,c_calibr_grav,c_calibr_fric,roughness_m,q_gas_sm3day)

    def calc_MF_calibr_pipe_m3day(self, qliq_sm3day,fw_perc,length_m,pin_atma,pout_atma,str_PVT=PVT_DEFAULT,theta_deg=90,d_mm=60,hydr_corr=H_CORRELATION,Tin_C=50,Tout_C=-1,c_calibr_grav=1,c_calibr_fric=1,roughness_m=0.0001,calibr_type=0):
        self.MF_calibr_pipe_m3day = self.book.macro("MF_calibr_pipe_m3day")
        return self.MF_calibr_pipe_m3day(qliq_sm3day,fw_perc,length_m,pin_atma,pout_atma,str_PVT,theta_deg,d_mm,hydr_corr,Tin_C,Tout_C,c_calibr_grav,c_calibr_fric,roughness_m,calibr_type)

    def calc_MF_p_pipe_atma(self, qliq_sm3day,fw_perc,length_m,pcalc_atma,calc_along_flow,str_PVT=PVT_DEFAULT,theta_deg=90,d_mm=60,hydr_corr=H_CORRELATION,tcalc_C=50,tother_C=-1,c_calibr_grav=1,c_calibr_fric=1,roughness_m=0.0001,q_gas_sm3day=0):
        self.MF_p_pipe_atma = self.book.macro("MF_p_pipe_atma")
        return self.MF_p_pipe_atma(qliq_sm3day,fw_perc,length_m,pcalc_atma,calc_along_flow,str_PVT,theta_deg,d_mm,hydr_corr,tcalc_C,tother_C,c_calibr_grav,c_calibr_fric,roughness_m,q_gas_sm3day)

    def calc_MF_p_pipe_znlf_atma(self, qliq_sm3day,fw_perc,length_m,pcalc_atma,calc_along_flow,str_PVT=PVT_DEFAULT,theta_deg=90,d_mm=60,hydr_corr=H_CORRELATION,tcalc_C=50,tother_C=-1,c_calibr_grav=1,c_calibr_fric=1,roughness_m=0.0001,Qgcas_free_scm3day=50):
        self.MF_p_pipe_znlf_atma = self.book.macro("MF_p_pipe_znlf_atma")
        return self.MF_p_pipe_znlf_atma(qliq_sm3day,fw_perc,length_m,pcalc_atma,calc_along_flow,str_PVT,theta_deg,d_mm,hydr_corr,tcalc_C,tother_C,c_calibr_grav,c_calibr_fric,roughness_m,Qgcas_free_scm3day)

    def calc_MF_dp_choke_atm(self, qliq_sm3day,fw_perc,dchoke_mm,pcalc_atma=-1,calc_along_flow=True,d_pipe_mm=70,tchoke_C=20,c_calibr_fr=1,str_PVT=PVT_DEFAULT):
        self.MF_dp_choke_atm = self.book.macro("MF_dp_choke_atm")
        return self.MF_dp_choke_atm(qliq_sm3day,fw_perc,dchoke_mm,pcalc_atma,calc_along_flow,d_pipe_mm,tchoke_C,c_calibr_fr,str_PVT)

    def calc_MF_p_choke_atma(self, qliq_sm3day,fw_perc,dchoke_mm,pcalc_atma=-1,calc_along_flow=True,d_pipe_mm=70,tchoke_C=20,c_calibr_fr=1,str_PVT=PVT_DEFAULT):
        self.MF_p_choke_atma = self.book.macro("MF_p_choke_atma")
        return self.MF_p_choke_atma(qliq_sm3day,fw_perc,dchoke_mm,pcalc_atma,calc_along_flow,d_pipe_mm,tchoke_C,c_calibr_fr,str_PVT)

    def calc_MF_calibr_choke_fr(self, qliq_sm3day,fw_perc,dchoke_mm,p_in_atma=-1,p_out_atma=-1,d_pipe_mm=70,tchoke_C=20,str_PVT=PVT_DEFAULT):
        self.MF_calibr_choke_fr = self.book.macro("MF_calibr_choke_fr")
        return self.MF_calibr_choke_fr(qliq_sm3day,fw_perc,dchoke_mm,p_in_atma,p_out_atma,d_pipe_mm,tchoke_C,str_PVT)

    def calc_MF_qliq_choke_sm3day(self, fw_perc,dchoke_mm,p_in_atma,p_out_atma,d_pipe_mm=70,tchoke_C=20,c_calibr_fr=1,str_PVT=PVT_DEFAULT):
        self.MF_qliq_choke_sm3day = self.book.macro("MF_qliq_choke_sm3day")
        return self.MF_qliq_choke_sm3day(fw_perc,dchoke_mm,p_in_atma,p_out_atma,d_pipe_mm,tchoke_C,c_calibr_fr,str_PVT)

    def calc_IPR_qliq_sm3day(self, pi_sm3dayatm,Pres_atma,pwf_atma,fw_perc=0,pb_atma=-1):
        self.IPR_qliq_sm3day = self.book.macro("IPR_qliq_sm3day")
        return self.IPR_qliq_sm3day(pi_sm3dayatm,Pres_atma,pwf_atma,fw_perc,pb_atma)

    def calc_IPR_pwf_atma(self, pi_sm3dayatm,Pres_atma,qliq_sm3day,fw_perc=0,pb_atma=-1):
        self.IPR_pwf_atma = self.book.macro("IPR_pwf_atma")
        return self.IPR_pwf_atma(pi_sm3dayatm,Pres_atma,qliq_sm3day,fw_perc,pb_atma)

    def calc_IPR_pi_sm3dayatm(self, Qtest_sm3day,Pwftest_atma,Pres_atma,fw_perc=0,pb_atma=-1):
        self.IPR_pi_sm3dayatm = self.book.macro("IPR_pi_sm3dayatm")
        return self.IPR_pi_sm3dayatm(Qtest_sm3day,Pwftest_atma,Pres_atma,fw_perc,pb_atma)

    def calc_ESP_head_m(self, qliq_m3day,num_stages=1,freq_Hz=50,pump_id=674,mu_cSt=-1,c_calibr_head=1,c_calibr_rate=1,c_calibr_power=1):
        self.ESP_head_m = self.book.macro("ESP_head_m")
        return self.ESP_head_m(qliq_m3day,num_stages,freq_Hz,pump_id,mu_cSt,c_calibr_head,c_calibr_rate,c_calibr_power)

    def calc_ESP_power_W(self, qliq_m3day,num_stages=1,freq_Hz=50,pump_id=674,mu_cSt=-1,c_calibr_rate=1,c_calibr_power=1):
        self.ESP_power_W = self.book.macro("ESP_power_W")
        return self.ESP_power_W(qliq_m3day,num_stages,freq_Hz,pump_id,mu_cSt,c_calibr_rate,c_calibr_power)

    def calc_ESP_eff_fr(self, qliq_m3day,num_stages=1,freq_Hz=50,pump_id=674,mu_cSt=-1,c_calibr_head=1,c_calibr_rate=1,c_calibr_power=1):
        self.ESP_eff_fr = self.book.macro("ESP_eff_fr")
        return self.ESP_eff_fr(qliq_m3day,num_stages,freq_Hz,pump_id,mu_cSt,c_calibr_head,c_calibr_rate,c_calibr_power)

    def calc_ESP_name(self, pump_id=674):
        self.ESP_name = self.book.macro("ESP_name")
        return self.ESP_name(pump_id)

    def calc_ESP_max_rate_m3day(self, freq_Hz=50,pump_id=674):
        self.ESP_max_rate_m3day = self.book.macro("ESP_max_rate_m3day")
        return self.ESP_max_rate_m3day(freq_Hz,pump_id)

    def calc_ESP_optRate_m3day(self, freq_Hz=50,pump_id=674):
        self.ESP_optRate_m3day = self.book.macro("ESP_optRate_m3day")
        return self.ESP_optRate_m3day(freq_Hz,pump_id)

    def calc_ESP_id_by_rate(self, q):
        self.ESP_id_by_rate = self.book.macro("ESP_id_by_rate")
        return self.ESP_id_by_rate(q)

    def calc_ESP_p_atma(self, qliq_sm3day,fw_perc,pcalc_atma,num_stages=1,freq_Hz=50,pump_id=674,str_PVT=PVT_DEFAULT,t_intake_C=50,t_dis_C=50,calc_along_flow=1,ESP_gas_degradation_type=0,c_calibr_head=1,c_calibr_rate=1,c_calibr_power=1):
        self.ESP_p_atma = self.book.macro("ESP_p_atma")
        return self.ESP_p_atma(qliq_sm3day,fw_perc,pcalc_atma,num_stages,freq_Hz,pump_id,str_PVT,t_intake_C,t_dis_C,calc_along_flow,ESP_gas_degradation_type,c_calibr_head,c_calibr_rate,c_calibr_power)

    def calc_ESP_dp_atm(self, qliq_sm3day,fw_perc,pcalc_atma,num_stages=1,freq_Hz=50,pump_id=674,str_PVT=PVT_DEFAULT,t_intake_C=50,t_dis_C=50,calc_along_flow=1,ESP_gas_degradation_type=0,c_calibr_head=1,c_calibr_rate=1,c_calibr_power=1):
        self.ESP_dp_atm = self.book.macro("ESP_dp_atm")
        return self.ESP_dp_atm(qliq_sm3day,fw_perc,pcalc_atma,num_stages,freq_Hz,pump_id,str_PVT,t_intake_C,t_dis_C,calc_along_flow,ESP_gas_degradation_type,c_calibr_head,c_calibr_rate,c_calibr_power)

    def calc_ESP_calibr_calc(self, qliq_sm3day,fw_perc,p_intake_atma,p_discharge_atma,str_PVT,str_ESP):
        self.ESP_calibr_calc = self.book.macro("ESP_calibr_calc")
        return self.ESP_calibr_calc(qliq_sm3day,fw_perc,p_intake_atma,p_discharge_atma,str_PVT,str_ESP)

    def calc_ESP_system_calc(self, qliq_sm3day,fw_perc,pcalc_atma,str_PVT,str_ESP,calc_along_flow=1):
        self.ESP_system_calc = self.book.macro("ESP_system_calc")
        return self.ESP_system_calc(qliq_sm3day,fw_perc,pcalc_atma,str_PVT,str_ESP,calc_along_flow)

    def calc_motor_M_slip_Nm(self, S,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_M_slip_Nm = self.book.macro("motor_M_slip_Nm")
        return self.motor_M_slip_Nm(S,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_I_slip_A(self, S,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_I_slip_A = self.book.macro("motor_I_slip_A")
        return self.motor_I_slip_A(S,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_CosPhi_slip(self, S,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_CosPhi_slip = self.book.macro("motor_CosPhi_slip")
        return self.motor_CosPhi_slip(S,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_Eff_slip(self, S,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_Eff_slip = self.book.macro("motor_Eff_slip")
        return self.motor_Eff_slip(S,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_M_Nm(self, Pshaft_kW,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_M_Nm = self.book.macro("motor_M_Nm")
        return self.motor_M_Nm(Pshaft_kW,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_I_A(self, Pshaft_kW,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_I_A = self.book.macro("motor_I_A")
        return self.motor_I_A(Pshaft_kW,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_CosPhi_d(self, Pshaft_kW,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_CosPhi_d = self.book.macro("motor_CosPhi_d")
        return self.motor_CosPhi_d(Pshaft_kW,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_Eff_d(self, Pshaft_kW,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_Eff_d = self.book.macro("motor_Eff_d")
        return self.motor_Eff_d(Pshaft_kW,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_S_d(self, Pshaft_kW,freq_Hz=50,U_V=-1,Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_S_d = self.book.macro("motor_S_d")
        return self.motor_S_d(Pshaft_kW,freq_Hz,U_V,Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_Name(self, Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_Name = self.book.macro("motor_Name")
        return self.motor_Name(Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_motor_Pnom_kW(self, Unom_V=500,Inom_A=10,Fnom_Hz=50,motorID=0):
        self.motor_Pnom_kW = self.book.macro("motor_Pnom_kW")
        return self.motor_Pnom_kW(Unom_V,Inom_A,Fnom_Hz,motorID)

    def calc_GLV_q_gas_sm3day(self, d_mm,p_in_atma,p_out_atma,gamma_g,t_C):
        self.GLV_q_gas_sm3day = self.book.macro("GLV_q_gas_sm3day")
        return self.GLV_q_gas_sm3day(d_mm,p_in_atma,p_out_atma,gamma_g,t_C)

    def calc_GLV_q_gas_vkr_sm3day(self, d_port_mm,d_vkr_mm,p_in_atma,p_out_atma,gamma_g,t_C):
        self.GLV_q_gas_vkr_sm3day = self.book.macro("GLV_q_gas_vkr_sm3day")
        return self.GLV_q_gas_vkr_sm3day(d_port_mm,d_vkr_mm,p_in_atma,p_out_atma,gamma_g,t_C)

    def calc_GLV_p_vkr_atma(self, d_port_mm,d_vkr_mm,p_calc_atma,q_gas_sm3day,gamma_g=0.6,t_C=25,calc_alog_flow=False):
        self.GLV_p_vkr_atma = self.book.macro("GLV_p_vkr_atma")
        return self.GLV_p_vkr_atma(d_port_mm,d_vkr_mm,p_calc_atma,q_gas_sm3day,gamma_g,t_C,calc_alog_flow)

    def calc_GLV_p_atma(self, d_mm,p_calc_atma,q_gas_sm3day,gamma_g=0.6,t_C=25,calc_alog_flow=False,p_open_atma=0):
        self.GLV_p_atma = self.book.macro("GLV_p_atma")
        return self.GLV_p_atma(d_mm,p_calc_atma,q_gas_sm3day,gamma_g,t_C,calc_alog_flow,p_open_atma)

    def calc_GLV_p_bellow_atma(self, p_atma,t_C):
        self.GLV_p_bellow_atma = self.book.macro("GLV_p_bellow_atma")
        return self.GLV_p_bellow_atma(p_atma,t_C)

    def calc_GLV_p_close_atma(self, p_bellow_atm,t_C):
        self.GLV_p_close_atma = self.book.macro("GLV_p_close_atma")
        return self.GLV_p_close_atma(p_bellow_atm,t_C)

    def calc_GLV_d_choke_mm(self, q_gas_sm3day,p_in_atma,p_out_atma,gamma_g=0.6,t_C=25):
        self.GLV_d_choke_mm = self.book.macro("GLV_d_choke_mm")
        return self.GLV_d_choke_mm(q_gas_sm3day,p_in_atma,p_out_atma,gamma_g,t_C)

    def calc_GLV_IPO_p_open(self, p_bellow_atma,p_out_atma,t_C,d_port_mm,GLV_type=0,d_vkr1_mm=-1,d_vkr2_mm=-1,d_vkr3_mm=-1,d_vkr4_mm=-1):
        self.GLV_IPO_p_open = self.book.macro("GLV_IPO_p_open")
        return self.GLV_IPO_p_open(p_bellow_atma,p_out_atma,t_C,GLV_type,d_port_mm,d_vkr1_mm,d_vkr2_mm,d_vkr3_mm,d_vkr4_mm)

    def calc_GLV_IPO_p_atma(self, p_bellow_atma,d_port_mm,p_calc_atma,q_gas_sm3day,t_C,calc_alog_flow=False,GLV_type=0,d_vkr1_mm=-1,d_vkr2_mm=-1,d_vkr3_mm=-1,d_vkr4_mm=-1):
        self.GLV_IPO_p_atma = self.book.macro("GLV_IPO_p_atma")
        return self.GLV_IPO_p_atma(p_bellow_atma,d_port_mm,p_calc_atma,q_gas_sm3day,t_C,calc_alog_flow,GLV_type,d_vkr1_mm,d_vkr2_mm,d_vkr3_mm,d_vkr4_mm)

    def calc_GLV_IPO_p_close(self, p_bellow_atma,p_out_atma,t_C,d_port_mm,GLV_type=0,d_vkr1_mm=-1,d_vkr2_mm=-1,d_vkr3_mm=-1,d_vkr4_mm=-1):
        self.GLV_IPO_p_close = self.book.macro("GLV_IPO_p_close")
        return self.GLV_IPO_p_close(p_bellow_atma,p_out_atma,t_C,GLV_type,d_port_mm,d_vkr1_mm,d_vkr2_mm,d_vkr3_mm,d_vkr4_mm)

    def calc_PVT_encode_string(self, gamma_gas=const_gg_,gamma_oil=const_go_,gamma_wat=const_gw_,rsb_m3m3=const_rsb_default,rp_m3m3=-1,pb_atma=-1,tres_C=const_tres_default,bob_m3m3=-1,muob_cP=-1,PVTcorr=Standing_based,ksep_fr=0,pksep_atma=-1,tksep_C=-1,gas_only=False):
        self.PVT_encode_string = self.book.macro("PVT_encode_string")
        return self.PVT_encode_string(gamma_gas,gamma_oil,gamma_wat,rsb_m3m3,rp_m3m3,pb_atma,tres_C,bob_m3m3,muob_cP,PVTcorr,ksep_fr,pksep_atma,tksep_C,gas_only)

    def calc_PVT_decode_string(self, str_PVT=PVT_DEFAULT,getStr=False):
        self.PVT_decode_string = self.book.macro("PVT_decode_string")
        return self.PVT_decode_string(str_PVT,getStr)

    def calc_well_encode_string(self, hperf_m=2000,hpump_m=1800,udl_m=0,d_cas_mm=150,dtub_mm=72,dchoke_mm=15,roughness_m=0.0001,tbh_C=85,twh_C=25):
        self.well_encode_string = self.book.macro("well_encode_string")
        return self.well_encode_string(hperf_m,hpump_m,udl_m,d_cas_mm,dtub_mm,dchoke_mm,roughness_m,tbh_C,twh_C)

    def calc_well_decode_string(self, str_well,getStr=False):
        self.well_decode_string = self.book.macro("well_decode_string")
        return self.well_decode_string(str_well,getStr)

    def calc_ESP_encode_string(self, esp_ID=1005,HeadNom_m=2000,ESPfreq_Hz=50,ESP_U_V=1000,MotorPowerNom_kW=30,t_intake_C=85,t_dis_C=85,KsepGS_fr=0,ESP_energy_fact_Whday=0,ESP_cable_type=0,ESP_Hmes_m=0,ESP_gas_degradation_type=0,c_calibr_head=0,c_calibr_rate=0,c_calibr_power=0,PKV_work_min=-1,PKV_stop_min=-1):
        self.ESP_encode_string = self.book.macro("ESP_encode_string")
        return self.ESP_encode_string(esp_ID,HeadNom_m,ESPfreq_Hz,ESP_U_V,MotorPowerNom_kW,t_intake_C,t_dis_C,KsepGS_fr,ESP_energy_fact_Whday,ESP_cable_type,ESP_Hmes_m,ESP_gas_degradation_type,c_calibr_head,c_calibr_rate,c_calibr_power,PKV_work_min,PKV_stop_min)

    def calc_ESP_decode_string(self, str_ESP,getStr=False):
        self.ESP_decode_string = self.book.macro("ESP_decode_string")
        return self.ESP_decode_string(str_ESP,getStr)

    def calc_wellGL_decode_string(self, well_GL_str,getStr=False):
        self.wellGL_decode_string = self.book.macro("wellGL_decode_string")
        return self.wellGL_decode_string(well_GL_str,getStr)

    def calc_wellGL_encode_string(self, hperf_m=2000,htub_m=1800,udl_m=0,d_cas_mm=150,dtub_mm=72,dchoke_mm=15,roughness_m=0.0001,tbh_C=85,twh_C=25,HmesGLV_m=0,dGLV_mm=0,PsurfGLV_atma=0):
        self.wellGL_encode_string = self.book.macro("wellGL_encode_string")
        return self.wellGL_encode_string(hperf_m,htub_m,udl_m,d_cas_mm,dtub_mm,dchoke_mm,roughness_m,tbh_C,twh_C,HmesGLV_m,dGLV_mm,PsurfGLV_atma)

    def calc_well_plin_pwf_atma(self, q_m3day,fw_perc,pwf_atma,Pcas_atma=-1,str_well=WELL_DEFAULT,str_PVT=PVT_DEFAULT,str_ESP=0,hydr_corr=H_CORRELATION,ksep_fr=0,c_calibr_head_d=0,param_num=1):
        self.well_plin_pwf_atma = self.book.macro("well_plin_pwf_atma")
        return self.well_plin_pwf_atma(q_m3day,fw_perc,pwf_atma,Pcas_atma,str_well,str_PVT,str_ESP,hydr_corr,ksep_fr,c_calibr_head_d,param_num)

    def calc_well_pintake_pwf_atma(self, q_m3day,fw_perc,pwf_atma,str_well=WELL_DEFAULT,str_PVT=PVT_DEFAULT,hydr_corr=H_CORRELATION):
        self.well_pintake_pwf_atma = self.book.macro("well_pintake_pwf_atma")
        return self.well_pintake_pwf_atma(q_m3day,fw_perc,pwf_atma,str_well,str_PVT,hydr_corr)

    def calc_well_pwf_plin_atma(self, q_m3day,fw_perc,plin_atma,Pcas_atma=-1,str_well=WELL_DEFAULT,str_PVT=PVT_DEFAULT,str_ESP=0,hydr_corr=H_CORRELATION,ksep_fr=-1,Psep_atma=40,Tsep_C=40,c_calibr_head_d=0,param_num=5):
        self.well_pwf_plin_atma = self.book.macro("well_pwf_plin_atma")
        return self.well_pwf_plin_atma(q_m3day,fw_perc,plin_atma,Pcas_atma,str_well,str_PVT,str_ESP,hydr_corr,ksep_fr,Psep_atma,Tsep_C,c_calibr_head_d,param_num)

    def calc_well_calcc_calibr_head_fr(self, q_m3day,fw_perc,Pdown_atma,pbuf_atma,Pdown_at_intake=False,plin_atma=-1,Pcas_atma=-1,str_well=WELL_DEFAULT,str_PVT=PVT_DEFAULT,str_ESP=0,hydr_corr=H_CORRELATION,ksep_fr=-1,c_calibr_head_d=0,param_num=0):
        self.well_calcc_calibr_head_fr = self.book.macro("well_calcc_calibr_head_fr")
        return self.well_calcc_calibr_head_fr(q_m3day,fw_perc,Pdown_atma,pbuf_atma,Pdown_at_intake,plin_atma,Pcas_atma,str_well,str_PVT,str_ESP,hydr_corr,ksep_fr,c_calibr_head_d,param_num)

    def calc_well_pwf_Hdyn_atma(self, q_m3day,fw_perc,Pcas_atma,Hdyn_m,str_well=WELL_DEFAULT,str_PVT=PVT_DEFAULT,str_ESP=0,hydr_corr=H_CORRELATION,ksep_fr=0,c_calibr_head_d=0,param_num=0):
        self.well_pwf_Hdyn_atma = self.book.macro("well_pwf_Hdyn_atma")
        return self.well_pwf_Hdyn_atma(q_m3day,fw_perc,Pcas_atma,Hdyn_m,str_well,str_PVT,str_ESP,hydr_corr,ksep_fr,c_calibr_head_d,param_num)

    def calc_nodal_qliq_sm3day(self, pi_sm3dayatm,plin_atma,fw_perc,Pres_atma=250,Pcas_atma=10,str_well=WELL_DEFAULT,str_PVT=PVT_DEFAULT,str_ESP=0,hydr_corr=H_CORRELATION,ksep_fr=0,c_calibr_head_d=0,param_num=23):
        self.nodal_qliq_sm3day = self.book.macro("nodal_qliq_sm3day")
        return self.nodal_qliq_sm3day(pi_sm3dayatm,plin_atma,fw_perc,Pres_atma,Pcas_atma,str_well,str_PVT,str_ESP,hydr_corr,ksep_fr,c_calibr_head_d,param_num)

    def calc_crv_interpolation(self, x_points,y_points,x_val,type_interpolation=0):
        self.crv_interpolation = self.book.macro("crv_interpolation")
        return self.crv_interpolation(x_points,y_points,x_val,type_interpolation)

    def calc_crv_solve(self, x_points,y_points,y_val):
        self.crv_solve = self.book.macro("crv_solve")
        return self.crv_solve(x_points,y_points,y_val)

    def calc_crv_intersection(self, x1_points,y1_points,x2_points,y2_points):
        self.crv_intersection = self.book.macro("crv_intersection")
        return self.crv_intersection(x1_points,y1_points,x2_points,y2_points)

UniflocVBA = API(addin_name_str)
print(UniflocVBA.calc_PVT_bg_m3m3(30,30))
