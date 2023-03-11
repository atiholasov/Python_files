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

old_etalon_data = pd.read_excel(path_etalon_file, header=None)
old_data_for_work = pd.read_excel(path_work_file, header=None)

    # Working with data

#Preparation work data
old_data_for_work.iloc[:, 0] = (round(old_data_for_work.iloc[:, 0]))
data_for_work = old_data_for_work

#Preparation etalon data
new_first_column = np.array(range(old_etalon_data.iloc[0, 0], old_etalon_data.iloc[-1, 0]+1))
len_new_first_column = len(new_first_column)
new_second_column = np.ones(len(new_first_column))
etalon_data = pd.DataFrame(np.column_stack([new_first_column, new_second_column]))
etalon_data = etalon_data.astype(int)

etalon_data_dict = {}
for index in old_etalon_data.index:
    etalon_data_dict[old_etalon_data.iloc[index, 0]] = old_etalon_data.iloc[index, 1]

for index in etalon_data.index:
    if etalon_data.iloc[index, 0] in etalon_data_dict:
        etalon_data.iloc[index, 1] = etalon_data_dict[etalon_data.iloc[index, 0]]
    else:
        etalon_data.iloc[index, 1] = np.nan

arr_concat = []
column_power = etalon_data.iloc[:,1]

for i in range(len(column_power)):
    if i == 0:
        list_for_interpolate = []
        list_for_interpolate.append(column_power[i])
        arr_concat.append(pd.Series(column_power[i]))
    else:
        list_for_interpolate.append(column_power[i])
        if not np.isnan(column_power[i]):
            list_for_interpolate = pd.Series(list_for_interpolate)
            list_for_interpolate = list_for_interpolate.interpolate()
            arr_concat.append(list_for_interpolate.iloc[1:])
            list_for_interpolate = [column_power[i]]
correct_column_of_power_in_etalon = pd.concat(arr_concat)
correct_column_of_power_in_etalon.index = range(len(correct_column_of_power_in_etalon))

etalon_data.iloc[:,1] = correct_column_of_power_in_etalon

# Reflectivity calculation

reflectivity = pd.DataFrame(np.zeros(2*len(data_for_work)).reshape(len(data_for_work), 2))  # final dataframe for plotting

# Нужно каждый элемент второго столбца у ворк дата по формуле смэтчить с подхрдящим по частоте элементом второго столбца 
# эталон дата

print(reflectivity.shape)
print(data_for_work.shape)
print(etalon_data.shape)