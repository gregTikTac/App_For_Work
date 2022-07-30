from pprint import pprint
from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
import openpyxl

doc = DocxTemplate('probe3.docx')
# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')
sheet = file_for_work.active

shortened_list_OG = []  # сокращенный список
full_list_OG = []  # полный список
list_of_month = {
    '1': 'января',
    '2': 'февраля',
    '3': 'марта',
    '4': 'апреля',
    '5': 'мая',
    '6': 'июня',
    '7': 'июля',
    '8': 'августа',
    '9': 'сентября',
    '10': 'октября',
    '11': 'ноября',
    '12': 'декабря'
}

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
window.geometry('600x400+500+200')

lbl = Label(window, text='Выберете месяц:', width=100, anchor=W)
lbl.pack()

def choose_month(event):
    print(select.current(), select.get())

select = ttk.Combobox(window, values=[
    'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
    'августа', 'сентября', 'октября', 'ноября', 'декабря'])
select.place(rely=0.08, anchor=W)
select.current(0)
select.bind("<<ComboboxSelected>>", choose_month)

window.mainloop()  # главный цикл