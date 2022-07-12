import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe2.docx')

# import database from exel
# data = pd.read_excel('word_automation.xlsm',
#                      skiprows=0,
#                      skipfooter=12,
#                      usecols='A',
#                     )

# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')

sheet = file_for_work.active

lst = []
for row in sheet.rows:
    lst.append(str(row[0].value))

del lst[0]
del lst[-12:]

dict_of_lst_og = {}
for index, value in enumerate(lst):
    dict_of_lst_og[index] = value
    # dict = {index: value for index, value in enumerate(lst} ---> то же самое
#
print(dict_of_lst_og)
sog = int(input("Выберите Старшего(нажмите соотвествующую цифру): "))



def change_sog_of_record():
    """Позволяет менять л/с в записи"""
    counter = 0
    while True:
        if sog == counter:
            context = {
                'sog': dict_of_lst_og[counter]
            }
            doc.render(context)
            doc.save('probe2.docx')
            break
        else:
            counter += 1


change_sog_of_record()