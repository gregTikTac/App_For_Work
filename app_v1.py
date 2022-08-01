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


def converting_exel_files_to_list_for_sog():
    """Перевод значений эксель в список (первый столбец)"""
    for row in sheet.rows:
        list_of_sog.append(str(row[0].value))


def converting_exel_files_to_list_for_og():
    """Перевод значений эксель в список (второй столбец столбец)"""
    for row in sheet.rows:
        list_of_og.append(str(row[1].value))


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
window.geometry('1200x700+150+050')
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


def choose_shorted_OG(event):
    print(select_SOG.get())
    print(select_OG)


converting_exel_files_to_list_for_sog()
converting_exel_files_to_list_for_og()
del list_of_og[0]
del list_of_sog[0]
del list_of_sog[-8:]

select_SOG = ttk.Combobox(window, values=list_of_sog, width=60)
select_SOG.place(rely=0.21, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_shorted_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.25, anchor=W)
select_OG.current(0)
select_OG.bind("<<ComboboxSelected>>", choose_shorted_OG)



# ОГ полного состава
lbl_OG_full = Label(window, text='Выберете ОГ полного состава:', anchor=W)
lbl_OG_full.place(rely=0.29)

def choose_full_OG(event):
    print(select_SOG.get())
    print(select_OG.get())
    print(select_OG.get())
    print(select_OG.get())
    print(select_OG.get())
    print(select_OG.get())
    print(select_OG.get())

select_SOG = ttk.Combobox(window, values=list_of_sog, width=60)
select_SOG.place(rely=0.33, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.37, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.41, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.45, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.49, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.53, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

select_OG = ttk.Combobox(window, values=list_of_og, width=60)
select_OG.place(rely=0.57, anchor=W)
select_SOG.current(0)
select_SOG.bind("<<ComboboxSelected>>", choose_full_OG)

window.mainloop()  # главный цикл
