# Created by Alexey Tiholasov (Tiholasov@yandex.ru)
# needed to: pip install openpyxl

import tkinter as tk
import numpy as np
import pandas as pd

# Graphics window
"""
path_work_file, path_etalon_file = "", ""


def func_with_out():
    global path_work_file, path_etalon_file
    path_work_file = work_file.get()
    path_etalon_file = etalon_file.get()
    win.quit()



# Work with open program window
win = tk.Tk()
win.title("Ploting spectral specular reflection coefficients")
win.geometry("500x200+730+300")
win.resizable(True, True)

text_for_etalon = tk.Label(win, text='Вставьте путь к эталонному файлу:')
text_for_data = tk.Label(win, text='Вставьте путь к измеренным данным:')
etalon_file = tk.Entry(win, width=60)
work_file = tk.Entry(win, width=60)

text_for_etalon.pack(pady=10)
etalon_file.pack(pady=5)
text_for_data.pack(pady=10)
work_file.pack(pady=10)

output_btn = tk.Button(win, text="Download data", command=func_with_out)
output_btn.pack(pady=10)

win.mainloop()
"""

# Reading files

#path_etalon_file = r"C:\Users\Ivan\Desktop\Alexey_Tiholasov\Script_for_specrofotometr\Data\2022-06-21_Al.xlsx"
#path_work_file = r"C:\Users\Ivan\Desktop\Alexey_Tiholasov\Script_for_specrofotometr\Data\Ag на кремнии.xlsx"

path_etalon_file = r"a_2022-06-21_Al.xlsx"
path_work_file = r"Ag на кремние.xlsx"

etalon_data = pd.read_excel(path_etalon_file, header=None)
data_for_work = pd.read_excel(path_work_file, header=None)

# Working with data

data_for_work.iloc[:, 0] = (round(data_for_work.iloc[:, 0]))

new_first_column = np.array(range(etalon_data.iloc[0, 0], etalon_data.iloc[-1, 0]+1))
len_new_first_column = len(new_first_column)
new_second_column = np.ones(len(new_first_column))
new_etalon_data = pd.DataFrame(np.column_stack([new_first_column, new_second_column]))
new_etalon_data = new_etalon_data.astype(int)

etalon_data_dict = {}
for index in etalon_data.index:
    etalon_data_dict[etalon_data.iloc[index, 0]] = etalon_data.iloc[index, 1]

for index in new_etalon_data.index:
    if new_etalon_data.iloc[index, 0] in etalon_data_dict:
        new_etalon_data.iloc[index, 1] = etalon_data_dict[new_etalon_data.iloc[index, 0]]
    else:
        new_etalon_data.iloc[index, 1] = np.nan

# new_etalon_data - хороший, но тупо пременить интерполяцию нельзя, она возьмет только первый и последний
# элементы. Нужно разбить на интервалы и их интерполировать, а затем склеить.
# Пробовал что то такое ниже


#print(new_etalon_data.head(31))

#new_etalon_data = new_etalon_data.iloc[:, 1].interpolate()
#print(new_etalon_data.head(31))


#Можно сделать функцию интерполяции интервала, чтобы было наглднее

a = [[95.1, np.nan, np.nan, 93.2],[93.2, np.nan, np.nan, 91.5]]
b = pd.Series(0)
for elem in a:
    aa = pd.Series(elem)
    aa = aa.interpolate()
    b = b.conco + aa
print(b)
