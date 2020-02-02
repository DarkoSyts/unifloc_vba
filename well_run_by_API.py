import description_generated.python_api as python_api
from scipy.optimize import minimize

UniflocVBA = python_api.API("UniflocVBA_7.xlam")
import sys

sys.path.append("../")


def calc_MF_q_mix_rc_m3day_for_minimize(qliq_sm3day, args):
    """
    Обертка исходной функции для minimize с функцией ошибки
    :param qliq_sm3day: подбираемый параметр- дебит на поверхности
    :param args: прочие аргументы для прямой расчетной функции
    :return: фуункция ошибки
    """
    qliq_sm3day = float(qliq_sm3day)
    q_mix_rc_m3day_true = args[0]
    fw_perc = args[1]
    p_atma = args[2]
    t_c = args[3]
    PVT_string = args[4]
    q_mix_rc_m3day = UniflocVBA.calc_MF_q_mix_rc_m3day(qliq_sm3day, fw_perc, p_atma, t_c, PVT_string)
    error = (q_mix_rc_m3day - q_mix_rc_m3day_true) ** 2
    return error


def q_liq_surface_m3day(q_mix_rc_m3day_true, fw_perc, p_atma, t_c, PVT_string):
    """
    Функция для расчета дебита на поверхности по дебиту в пластовых условиях
    :param q_mix_rc_m3day_true:
    :param fw_perc:
    :param p_atma:
    :param t_c:
    :param PVT_string:
    :return:
    """
    qliq_sm3day = 100
    minimize_result = minimize(calc_MF_q_mix_rc_m3day_for_minimize, qliq_sm3day, args=[q_mix_rc_m3day_true, fw_perc, p_atma, t_c, PVT_string])
    print("Сводная информация minimize")
    print(minimize_result)
    q_liq_surface_m3day = minimize_result.x[0]
    return q_liq_surface_m3day


PVT_string = UniflocVBA.calc_PVT_encode_string()
qliq_sm3day = 50
fw_perc = 50
p_atma = 20
t_c = 30

q_mix_rc_m3day_true = UniflocVBA.calc_MF_q_mix_rc_m3day(qliq_sm3day, fw_perc, p_atma, t_c, PVT_string)

q_liq_surface_m3day_calculated = q_liq_surface_m3day(q_mix_rc_m3day_true, fw_perc, p_atma, t_c, PVT_string)

print(f"Дебит в условиях насоса {q_mix_rc_m3day_true} при заданном значении дебита на поверхности {qliq_sm3day}")

print(f"Дебит в на поверхности подобранный {q_liq_surface_m3day_calculated} при известном значении дебита"
      f"в пластовых условиях {q_mix_rc_m3day_true}")
pvt_str = UniflocVBA.calc_PVT_encode_string(gamma_gas=1,
                                                       gamma_oil=1,
                                                       gamma_wat=1,
                                                       rsb_m3m3=1,
                                                       rp_m3m3=1,
                                                       pb_atma=10,
                                                       tres_C=20,
                                                       bob_m3m3=1,
                                                       muob_cP=1,
                                                       ksep_fr=0.2,
                                                       pksep_atma=1,
                                                       tksep_C=1)