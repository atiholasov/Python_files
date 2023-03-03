# Created by Alexey Tiholasov (Tiholasov@yandex.ru)
# needed to: pip install openpyxl

import tkinter as tk
import pandas as pd

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








# Working with data

#path_etalon_file = r"C:\Users\Ivan\Desktop\Alexey_Tiholasov\Script_for_specrofotometr\Data\2022-06-21_Al.xlsx"
#path_work_file = r"C:\Users\Ivan\Desktop\Alexey_Tiholasov\Script_for_specrofotometr\Data\Ag на кремнии.xlsx"

etalon_data = pd.read_excel(r"a_2022-06-21_Al.xlsx", header=None)
#data_for_work = pd.read_excel(path_work_file, header=None)

print(etalon_data)
#print(data_for_work)