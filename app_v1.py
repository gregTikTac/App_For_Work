from pprint import pprint
from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe3.docx')
# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')
sheet = file_for_work.active

list_of_sog = []  # сокращенный список
list_of_og = []  # полный список
list_of_pat = []
list_PM = []
list_AK = []


def converting_exel_files_to_list_for_sog():
    """Перевод значений эксель в список (первый столбец)"""
    for row in sheet.rows:
        list_of_sog.append(str(row[0].value))


def converting_exel_files_to_list_for_og():
    """Перевод значений эксель в список (второй столбец столбец)"""
    for row in sheet.rows:
        list_of_og.append(str(row[1].value))


def converting_exel_files_to_list_for_PAT():
    """Конвертирует данные из 3 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_of_pat.append(str(row[2].value))


def converting_exel_files_to_list_PM_for_OG():
    """Конвертирует данные из 4 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_PM.append(str(row[3].value))


def converting_exel_files_to_list_AK_for_PAT():
    """Конвертирует данные из 5 столбца таблицы exel в список"""
    for row in sheet.rows:
        list_AK.append(str(row[4].value))
#
# def print_all_month():
#     for key, value in list_of_month.items():
#         print(f'  {key} - {value}')
#     print('+' + '_ _ _' * 15 + '+')
#
#
# def choose_num_month(num_month):
#     for key, value in enumerate(list_of_month):
#         print(key, value)
#
#
#
# print_all_month()
# choose_num_month(num_month=input("Введите номер месяца: "))
window = Tk()  # объект окна

window.title('GUI для работы')
window.geometry('1300x750+150+050')
window.resizable(False, False)
# label месяца
lbl_month = Label(window, text='Выберете месяц:', width=100, anchor=W)
lbl_month.place(rely=0.01)


def choose_month(event):
    """обработчик событий"""
    print(select_month.current(), select_month.get())


# выпадающее окно
select_month = ttk.Combobox(window, values=[
    'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
    'августа', 'сентября', 'октября', 'ноября', 'декабря'])
select_month.place(rely=0.05, anchor=W)
select_month.current(0)
select_month.bind("<<ComboboxSelected>>", choose_month)

# дата
lbl_date = Label(window, text='Выберете дату:', anchor=W)
lbl_date.place(rely=0.09)


def choose_date(event):
    print(select_date.get())


select_date = ttk.Combobox(window, values=[i for i in range(1, 32)])
select_date.place(rely=0.13, anchor=W)
select_date.current(0)
select_date.bind("<<ComboboxSelected>>", choose_date)



# ОГ сокращенного состава
lbl_OG_shorted = Label(window, text='Выберете ОГ сокращенного сотава:', anchor=W)
lbl_OG_shorted.place(rely=0.17)


converting_exel_files_to_list_for_sog()
converting_exel_files_to_list_for_og()
del list_of_og[0]
del list_of_sog[0]
del list_of_sog[-8:]

select_SOG = ttk.Combobox(window, values=list_of_sog, width=60)
select_SOG.place(rely=0.21, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.24, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")



# ОГ полного состава
lbl_OG_full = Label(window, text='Выберете ОГ полного состава:', anchor=W)
lbl_OG_full.place(rely=0.29)


select_SOG = ttk.Combobox(window, values=list_of_sog, width=60)
select_SOG.place(rely=0.33, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.36, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.39, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.421, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.45, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.481, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.512, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>")


# ПАТ
lbl_PAT = Label(window, text='Выберете ПАТ:', anchor=W)
lbl_PAT.place(relx=0.35, rely=0.01)


converting_exel_files_to_list_for_PAT()
del list_of_pat[0]
select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.04)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.07)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.10)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.13)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.16)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.19)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.22)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.25)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.28)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.31)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.34)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.37)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.40)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.43)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.460)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.491)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.52)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.55)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.58)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.61)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.64)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.67)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.70)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.73)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")

select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
select_PAT.place(relx=0.35, rely=0.76)
select_PAT.current(0)
select_SOG.bind("<<ComboboxSelected>>")


#ПМ
converting_exel_files_to_list_PM_for_OG()
del list_PM[0]
lbl_PM_shorted_OG = Label(window, text='ПМ для ОГ сокращенного состава:', anchor=W)
lbl_PM_shorted_OG.place(rely=0.56)

select_PM = ttk.Combobox(window, values=list_PM, width=40)
select_PM.place(rely=0.585)

select_PM = ttk.Combobox(window, values=list_PM, width=40)
select_PM.place(rely=0.615)



lbl_PM_full_OG = Label(window, text='ПМ для ОГ полного состава:', anchor=W)
lbl_PM_full_OG.place(rely=0.65)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.675)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.705)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.735)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.765)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.795)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.825)

select_PM = ttk.Combobox(window, values=list_PM, width=60)
select_PM.place(rely=0.855)

window.mainloop()  # главный цикл
