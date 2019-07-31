import xlwings as xw
import pandas as pd
book = xw.Book("UF7_калькулятор_ЭЦН _1354_с_СУ.xlsm")

app = book.macro("get_data")

'pvt = book.macro("PVT_Bg_m3m3")'

check = app()

print(check)
check_dict = {}
check_dict.update({'Дата': [28.02]})
values = check[0]
print(values)
names = check[1]
for name, value in zip(names, values):
    check_dict.update({name: [value]})

df_check = pd.DataFrame(check_dict)
print(df_check)
df_check.to_csv("df_check2.csv")
