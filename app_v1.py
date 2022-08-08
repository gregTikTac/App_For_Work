import tkinter as tk
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
list_for_render = []


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


converting_exel_files_to_list_PM_for_OG()
converting_exel_files_to_list_for_sog()
converting_exel_files_to_list_for_og()
converting_exel_files_to_list_for_PAT()
converting_exel_files_to_list_AK_for_PAT()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('GUI for work')
        self.geometry('1200x700+150+050')

        self.variable = tk.StringVar()

        self.label_month = ttk.Label(text="Выберете месяц:").grid()

        self.combobox_month = ttk.Combobox(values=['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
                                                   'августа', 'сентября', 'октября', 'ноября', 'декабря'], width=40)
        self.combobox_month.grid()
        self.combobox_month.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_days = ttk.Label(text="Выберете день:").grid()
        self.combobox_days = ttk.Combobox(values=[i for i in range(1, 32)], width=40)
        self.combobox_days.grid()
        self.combobox_days.bind('<<ComboboxSelected>>', self.add_entry)

        # OG
        self.label_shorted_og = ttk.Label(text="Выберете ОГ сокращенного состава:").grid()
        self.combobox_sog = ttk.Combobox(values=list_of_sog, width=40)
        self.combobox_sog.grid()
        self.combobox_sog.bind('<<ComboboxSelected>>', self.add_entry)
        self.combobox_og = ttk.Combobox(values=list_of_og, width=40)
        self.combobox_og.grid()
        self.combobox_og.bind('<<ComboboxSelected>>', self.add_entry)

        # OG FULL
        self.label_full_squad = ttk.Label(text="Выберете ОГ полного состава:").grid()
        self.combobox_full_sog = ttk.Combobox(values=list_of_sog, width=40)
        self.combobox_full_sog.grid()
        self.combobox_full_sog.bind('<<ComboboxSelected>>', self.add_entry)
        for item in range(1, 7):
            self.com = ttk.Combobox(values=list_of_og, width=40)
            self.com.grid()
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        # PM
        self.label_shorted_PM = ttk.Label(text="Выберете PM для ОГ сокращенного состава:").grid()
        self.combobox_pm_for_shorted_og = ttk.Combobox(values=list_PM, width=40)
        self.combobox_pm_for_shorted_og.grid()
        self.combobox_pm_for_shorted_og.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_full_PM = ttk.Label(text="Выберете PM для ОГ полного состава:").grid()
        for item in range(1, 8):
            self.com = ttk.Combobox(values=list_PM, width=40)
            self.com.grid()
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        # PAT
        self.label_PAT = ttk.Label(text="Выберете PAT состава:").grid(row=0, column=2)
        self.label_control_group = ttk.Label(text="Control group").grid(row=1, column=2)
        for item in range(2, 4):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ranger_group = ttk.Label(text="ranger group").grid(row=4, column=2)
        for item in range(5, 7):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_defense_group = ttk.Label(text="defense group").grid(row=7, column=2)
        for item in range(8, 16):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_fire_group = ttk.Label(text="fire group").grid(row=16, column=2)
        for item in range(17, 20):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_reserve_group = ttk.Label(text="reserve group").grid(row=20, column=2)
        for item in range(21, 25):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_sappers_group = ttk.Label(text="sappers group").grid(row=25, column=2)
        for item in range(26, 28):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_rhb_group = ttk.Label(text="rhb group").grid(row=28, column=2)
        for item in range(29, 31):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_med_group = ttk.Label(text="med group").grid(row=31, column=2)
        for item in range(32, 34):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        # AK
        self.label_ak_PAT = ttk.Label(text="Выберете ak PAT :").grid(row=0, column=3)
        self.label_ak_control_group = ttk.Label(text="Control group").grid(row=1, column=3)
        for item in range(2, 4):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_ranger_group = ttk.Label(text="ranger group").grid(row=4, column=3)
        for item in range(5, 7):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_defense_group = ttk.Label(text="defense group").grid(row=7, column=3)
        for item in range(8, 16):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_fire_group = ttk.Label(text="fire group").grid(row=16, column=3)
        for item in range(17, 20):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_reserve_group = ttk.Label(text="reserve group").grid(row=20, column=3)
        for item in range(21, 25):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_sappers_group = ttk.Label(text="sappers group").grid(row=25, column=3)
        for item in range(26, 28):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_rhb_group = ttk.Label(text="rhb group").grid(row=28, column=3)
        for item in range(29, 31):
            self.com = ttk.Combobox(values=list_AK, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.label_ak_med_group = ttk.Label(text="med group").grid(row=31, column=3)
        for item in range(32, 34):
            self.com = ttk.Combobox(values=list_of_pat, width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry)

        self.btn = tk.Button(text='push me!', width=10, heigh=3)
        self.btn.grid(row=10, column=4)
        self.btn.bind('<Button-1>', self.push_btn)

    def push_btn(self, event):
        print(list_for_render)

    def add_entry(self, event):
        print(self.variable.get())
        if event:
            print(f'выбрано: {event.widget.get()}')
            list_for_render.append(event.widget.get())


app = App()
app.mainloop()
