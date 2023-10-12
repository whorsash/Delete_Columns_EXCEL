from tkinter import Tk

import tkinter.filedialog as fd

Tk().withdraw()
print('Выберите файл')
filename = fd.askopenfilename()
while filename == '':
    print('Не выбран файл EXCEL')
    print('Выберите файл')
    filename = fd.askopenfilename()

print('Выберите путь для сохранения нового файла')
directory = fd.askdirectory()
while directory == '':
    print('Не выбран путь для сохранения нового файла')
    print('Выберите путь для сохранения нового файла')
    directory = fd.askdirectory()

import pandas as pd

print(f'Файл {filename} читаем!')
df = pd.read_excel(filename)
print(f'Файл {filename} прочитан!')


print('Удаляем колонки!')
del df['НазваниеКолонки1']
del df['НазваниеКолонки2']
del df['НазваниеКолонки3']
del df['НазваниеКолонки4']
del df['НазваниеКолонки5']
del df['НазваниеКолонки6']
del df['НазваниеКолонки7']
del df['НазваниеКолонки8']

print('Колонки удалены!')
print('Записываем файл new_file.xlsx')
writer = pd.ExcelWriter(directory + '/new_file.xlsx', engine='xlsxwriter')

sheet = pd.ExcelFile(filename).sheet_names
print(sheet)
df.to_excel(writer, 'Sheet1')
writer.save()
print('Файл записан new_file.xlsx!')

import os

file_xlsx = directory + '/new_file.xlsx'

if os.access(file_xlsx, os.F_OK) != True:

    print("Файл не существует!")

else:
    print("Файл существует")

print("Читаем файл new_file.xlsx")
# Install the openpyxl library
from openpyxl import load_workbook

# Loading our Excel file
wb = load_workbook(file_xlsx)

# creating the sheet 1 object
ws = wb.worksheets[0]
print("Удаляем первую колонку")
# удалим 1 столбцы в диапазоне `F:H`
ws.delete_cols(1, 1)

#datetime
import datetime as dt
from pytz import timezone
import pytz
date = dt.datetime.now().strftime("%d_%m_%y")
print(date)
# сохраняемся и открываем файл
wb.save(file_xlsx)
wb.save(directory + '/new_file_new' + date + '_.xlsx')

print("Файл new_file_new.xlsx успешно записан!")
