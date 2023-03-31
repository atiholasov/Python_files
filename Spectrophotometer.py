# Created by Alexey Tiholasov (Tiholasov@yandex.ru)
# needed to: pip install openpyxl

import tkinter as tk
from tkinter import ttk, filedialog
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

path_work_file, path_etalon_file, name_et_file, name_wr_file = "", "", "", ""


def prepare_file_name(file_path):
    file_name = list(str(file_path).split("/"))[-1]
    return file_name

    # Graphics window


def func_with_out():
    win.quit()


def open_file_et():
    global path_etalon_file, name_et_file
    file = filedialog.askopenfile(mode='r', filetypes=[("Excel files", "*.xlsx")])
    if file:
       filepath = os.path.abspath(file.name)
       tk.Label(win, text="Эталонный файл : " + str(filepath)).pack(pady=10)
       path_etalon_file = filepath
       name_et_file = prepare_file_name(file.name)


def open_file_wr():
    global path_work_file, name_wr_file
    file = filedialog.askopenfile(mode='r', filetypes=[("Dat Files", "*.dat")])
    if file:
       filepath = os.path.abspath(file.name)
       tk.Label(win, text="Измеренный файл : " + str(filepath)).pack(pady=10)
       path_work_file = filepath
       name_wr_file = prepare_file_name(file.name)
       output_btn = tk.Button(win, text="Download data", command=func_with_out)
       output_btn.pack(pady=10)

# Work with open program window
win = tk.Tk()
win.title("Ploting spectral specular reflection coefficients")
win.geometry("550x230+730+300")
win.resizable(True, True)

Br_1 = ttk.Button(win, text="Выберите путь к эталонному файлу (.xlsx)", command=open_file_et)
Br_2 = ttk.Button(win, text="Выберите путь к измеренным данным (.dat)", command=open_file_wr)
Br_1.pack(pady=15)
Br_2.pack(pady=10)

win.mainloop()

old_etalon_data = pd.read_excel(path_etalon_file, header=None)
old_data_for_work = pd.read_table(path_work_file, header=None, sep='\t')

    # Working with data

#Preparation work data
old_data_for_work.iloc[:, 0] = (round(old_data_for_work.iloc[:, 0]))
data_for_work = old_data_for_work

#Preparation etalon data
new_first_column = np.array(range(old_etalon_data.iloc[0, 0], old_etalon_data.iloc[-1, 0]+1))
len_new_first_column = len(new_first_column)
new_second_column = np.ones(len(new_first_column))
etalon_data = pd.DataFrame(np.column_stack([new_first_column, new_second_column]))
etalon_data = etalon_data.astype(float)

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

etalon_data.iloc[:, 1] = correct_column_of_power_in_etalon

# Reflectivity calculation

reflectivity = pd.DataFrame(np.zeros(3*len(data_for_work)).reshape(len(data_for_work), 3))  # final dataframe for plotting
reflectivity.index = data_for_work.iloc[:,0].astype(float)
reflectivity.columns = ["etalon", "work_data", "true_data"]
data_for_work.index = reflectivity.index
etalon_data.index = etalon_data.iloc[:,0].astype(float)
etalon_data.columns = ["wavelength", "power"]
data_for_work.columns = ["wavelength", "power"]
reflectivity["work_data"] = data_for_work.iloc[:, 1]

list_for_delite = []   # delited wavelength in etalon data wich not in work_data 
for lenght in etalon_data.wavelength:
    if lenght not in data_for_work.wavelength:
        list_for_delite.append(lenght)

correct_et_data = etalon_data.drop(list_for_delite)

reflectivity["etalon"] = correct_et_data.iloc[:, 1]
reflectivity["true_data"] = reflectivity["etalon"] * reflectivity["work_data"] / 100  # true_data is (power_et * power_work / 100)
reflectivity_for_exel = pd.DataFrame({"Wavelength, [nm]": reflectivity.index,
                                      f"Reference values ({name_et_file})": reflectivity["etalon"],
                                      f"Measured values ({name_wr_file})": reflectivity["work_data"],
                                      "Real values": reflectivity["true_data"]})

path_for_out_file = path_work_file[:-len(name_wr_file)]
writer = pd.ExcelWriter(f"{path_for_out_file}Result for {name_wr_file[:-4]}.xlsx")
reflectivity_for_exel.to_excel(writer, index=False)
writer.close()

    # Plotting

plt.rcParams["figure.autolayout"] = True
plt.rcParams["figure.figsize"] = [15, 6]
plt.figure()
plt.title('Spectral reflection characteristic')
plt.xlabel("wavelength, [nm]")
plt.ylabel("reflective characteristic")
plt.minorticks_on()
plt.grid(which='major', color="gray")
plt.grid(which='minor', linestyle='dotted', axis="both", color="gray")
plt.plot(reflectivity["true_data"])
plt.show()