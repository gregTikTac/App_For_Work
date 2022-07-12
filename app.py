import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe.docx')

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

print(dict_of_lst_og)
sog = int(input("Выберите Старшего(нажмите соотвествующую цифру): "))

if sog == 0:
    context = {
        'sog': dict_of_lst_og[0]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 1:
    context = {
        'sog': dict_of_lst_og[1]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 2:
    context = {
        'sog': dict_of_lst_og[2]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 3:
    context = {
        'sog': dict_of_lst_og[3]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 4:
    context = {
        'sog': dict_of_lst_og[4]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 5:
    context = {
        'sog': dict_of_lst_og[5]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 6:
    context = {
        'sog': dict_of_lst_og[6]
    }
    doc.render(context)
    doc.save('probe.docx')

elif sog == 7:
    context = {
        'sog': dict_of_lst_og[7]
    }
    doc.render(context)
    doc.save('probe.docx')

else:
    print('Такого значения нет')

