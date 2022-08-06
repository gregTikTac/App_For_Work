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


window = Tk()  # объект окна

window.title('GUI для работы')
window.geometry('1200x700+150+050')


# window.resizable(False, False)


class Block:
    """Блок лейбл и комбобос виджетов"""

    def __init__(self, window, label, combobox):
        self.window = window
        self.label = label
        self.combobox = combobox


# месяц
select_month = Block(window, label=Label(text="Выберете месяц: ").grid(),
                     combobox=ttk.Combobox(
                         values=['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
                                 'августа', 'сентября', 'октября', 'ноября', 'декабря'],
                         width=40).grid())

# дата
select_date = Block(window, label=Label(text="Выберете дату:").grid(),
                    combobox=ttk.Combobox(values=[i for i in range(1, 32)], width=40).grid())

converting_exel_files_to_list_for_sog()
converting_exel_files_to_list_for_og()

# ОГ сокращенного состава
select_shorted_SOG = Block(window, label=Label(text='Выберете CОГ сокращенного сотава:').grid(),
                           combobox=ttk.Combobox(values=list_of_sog, width=60).grid())

select_shorted_OG = Block(window, label=None, combobox=ttk.Combobox(values=list_of_og, width=60).grid())

# ОГ полного состава
select_full_squad = Block(window, label=Label(text='Выберете ОГ полного сотава:').grid(),
                          combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_full_squad2 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_og, width=60).grid())
select_full_squad3 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_og, width=60).grid())
select_full_squad4 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_og, width=60).grid())
select_full_squad5 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_og, width=60).grid())
select_full_squad6 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_og, width=60).grid())
#
# ПМ
converting_exel_files_to_list_PM_for_OG()
select_pm_for_shorted_og1 = Block(window, label=Label(window, text="Выберете PM for shorted squad:").grid(),
                                  combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_shorted_og2 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())

select_pm_for_full_og1 = Block(window, label=Label(window, text="Выберете PM for full squad:").grid(),
                               combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_full_og2 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_full_og3 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_full_og4 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_full_og5 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_full_og6 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())
select_pm_for_full_og7 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_sog, width=60).grid())

# ПАТ
converting_exel_files_to_list_for_PAT()


select_pat1 = Block(window, label=Label(window, text="ГР. УПР.:").grid(row=1, column=2),
                    combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=2, column=2))
select_pat2 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=3, column=2))


select_pat3 = Block(window, label=Label(window, text="Гр. Рзвд.").grid(row=4, column=2),
                    combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=5, column=2))
select_pat4 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=6, column=2))


select_pat5 = Block(window, label=Label(window, text="Гр. БЛК.").grid(row=7, column=2),
                    combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=8, column=2))
select_pat6 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=9, column=2))
select_pat7 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=10, column=2))
select_pat8 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=11, column=2))
select_pat9 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=12, column=2))
select_pat10 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=13, column=2))
select_pat11 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=14, column=2))
select_pat12 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=15, column=2))


select_pat13 = Block(window, label=Label(window, text="Гр. Огн.П.").grid(row=16, column=2),
                     combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=17, column=2))
select_pat14 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=18, column=2))
select_pat15 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=19, column=2))


select_pat16 = Block(window, label=Label(window, text="Гр. РЗВР.").grid(row=20, column=2),
                     combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=21, column=2))
select_pat17 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=22, column=2))
select_pat18 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=23, column=2))
select_pat19 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=24, column=2))


select_pat20 = Block(window, label=Label(window, text="ОТД. РЗМР.").grid(row=25, column=2),
                     combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=26, column=2))
select_pat21 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=27, column=2))

select_pat22 = Block(window, label=Label(window, text="ОТД. РХБЗ.").grid(row=25, column=2),
                     combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=26, column=2))
select_pat23 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=27, column=2))

select_pat24 = Block(window, label=Label(window, text="ОТД. МЕД.").grid(row=28, column=2),
                     combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=29, column=2))
select_pat25 = Block(window, label=None, combobox=ttk.Combobox(values=list_of_pat, width=60).grid(row=30, column=2))
#
#


# def choose_month(event):
#     """обработчик событий"""
#     print(select_month.current(), select_month.get())
#
#
# # выпадающее окно
# select_month = SelectComboBox(ttk.Combobox(window, values=[
#     'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
#     'августа', 'сентября', 'октября', 'ноября', 'декабря']).place(rely=0.05, anchor=W)
# # select_month.pla()
# # select_month.current(0)
# # select_month.bind("<<ComboboxSelected>>", choose_month)
#
# # дата
# lbl_date = Label(window, text='Выберете дату:', anchor=W)
# lbl_date.place(rely=0.09)
#
#
# def choose_date(event):
#     print(select_date.get())
#
#
# select_date = ttk.Combobox(window, values=[i for i in range(1, 32)])
# select_date.place(rely=0.13, anchor=W)
# select_date.current(0)
# select_date.bind("<<ComboboxSelected>>", choose_date)
#
# # ОГ сокращенного состава
# lbl_OG_shorted = Label(window, text='Выберете ОГ сокращенного сотава:', anchor=W)
# lbl_OG_shorted.place(rely=0.17)
#
# converting_exel_files_to_list_for_sog()
# converting_exel_files_to_list_for_og()
# del list_of_og[0]
# del list_of_sog[0]
# del list_of_sog[-8:]
#
# select_SOG = ttk.Combobox(window, values=list_of_sog, width=60)
# select_SOG.place(rely=0.21, anchor=W)
# select_SOG.current(0)
# select_SOG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.24, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# # ОГ полного состава
# lbl_OG_full = Label(window, text='Выберете ОГ полного состава:', anchor=W)
# lbl_OG_full.place(rely=0.29)
#
# select_SOG = ttk.Combobox(window, values=list_of_sog, width=60)
# select_SOG.place(rely=0.33, anchor=W)
# select_SOG.current(0)
# select_SOG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.36, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.39, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.421, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.45, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.481, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# select_OG = ttk.Combobox(window, values=list_of_og, width=60)
# select_OG.place(rely=0.512, anchor=W)
# select_OG.current(0)
# select_OG.bind("<<ComboboxSelected>>")
#
# # ПАТ
# lbl_PAT = Label(window, text='Выберете ПАТ:', anchor=W)
# lbl_PAT.place(relx=0.35, rely=0.01)
#
# converting_exel_files_to_list_for_PAT()
# del list_of_pat[0]
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.04)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.07)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.10)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.13)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.16)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.19)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.22)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.25)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.28)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.31)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.34)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.37)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.40)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.43)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.460)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.491)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.52)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.55)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.58)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.61)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.64)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.67)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.70)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.73)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# select_PAT = ttk.Combobox(window, values=list_of_pat, width=40)
# select_PAT.place(relx=0.35, rely=0.76)
# select_PAT.current(0)
# select_PAT.bind("<<ComboboxSelected>>")
#
# # ПМ
# converting_exel_files_to_list_PM_for_OG()
# del list_PM[0]
# lbl_PM_shorted_OG = Label(window, text='ПМ для ОГ сокращенного состава:', anchor=W)
# lbl_PM_shorted_OG.place(rely=0.56)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=40)
# select_PM.place(rely=0.585)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=40)
# select_PM.place(rely=0.615)
#
# lbl_PM_full_OG = Label(window, text='ПМ для ОГ полного состава:', anchor=W)
# lbl_PM_full_OG.place(rely=0.65)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.675)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.705)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.735)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.765)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.795)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.825)
#
# select_PM = ttk.Combobox(window, values=list_PM, width=60)
# select_PM.place(rely=0.855)
#
# # АК для Пат
# converting_exel_files_to_list_AK_for_PAT()
# del list_AK[0]
# lbl_PAT = Label(window, text='АК для ПАТ:', anchor=W)
# lbl_PAT.place(relx=0.60, rely=0.01)
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.04)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.07)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.10)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.13)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.16)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.19)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.22)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.25)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.28)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.31)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.34)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.37)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.40)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.43)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.460)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.491)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.52)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.55)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.58)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.61)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.64)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.67)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.70)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.73)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
# select_AK = ttk.Combobox(window, values=list_AK, width=40)
# select_AK.place(relx=0.60, rely=0.76)
# select_AK.current(0)
# select_AK.bind("<<ComboboxSelected>>")
#
#
# object1 = ttk.Combobox(values=list_AK, text='proba', width=55, relx=0.66, rely=0.80)
window.mainloop()  # главный цикл
