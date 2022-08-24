import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as mb
from docxtpl import DocxTemplate
import openpyxl
from XlsLoader import XlsLoader
import Constants

doc = DocxTemplate('probe3.docx')

# читаем файл
file_for_work = openpyxl.load_workbook('word_automation.xlsm')

xls_loader = XlsLoader('word_automation.xlsm')
xls_loader.load()

sheet = file_for_work.active
list_of_sog = []
list_of_og = []
list_of_pat = []
list_PM = []
list_AK = []
list_for_render_date_and_month = []
list_for_render_shorted_og = []
list_for_render_full_og = []
list_for_render_shorted_og_pm = []
list_for_render_full_og_pm = []
list_for_render_pat = []
list_for_render_pat_ak = []


# def converting_exel_files_to_list_for_sog():
#     """Перевод значений эксель в список (первый столбец)"""
#     for row in sheet.rows:
#         list_of_sog.append(str(row[0].value))
#     del list_of_sog[0]
#     del list_of_sog[-32:]
#
#
# def converting_exel_files_to_list_for_og():
#     """Перевод значений эксель в список (второй столбец столбец)"""
#     for row in sheet.rows:
#         list_of_og.append(str(row[1].value))
#     del list_of_og[0]
#     del list_of_og[-20:]
#
#
# def converting_exel_files_to_list_for_PAT():
#     """Конвертирует данные из 3 столбца таблицы exel в список"""
#     for row in sheet.rows:
#         list_of_pat.append(str(row[2].value))
#     del list_of_pat[0]
#     del list_of_pat[-15:]
#
#
# def converting_exel_files_to_list_PM_for_OG():
#     """Конвертирует данные из 4 столбца таблицы exel в список"""
#     for row in sheet.rows:
#         list_PM.append(str(row[3].value))
#     del list_PM[0]
#     del list_PM[-2:]
#
#
# def converting_exel_files_to_list_AK_for_PAT():
#     """Конвертирует данные из 5 столбца таблицы exel в список"""
#     for row in sheet.rows:
#         list_AK.append(str(row[4].value))
#     del list_AK[-8:]
#     del list_AK[0]


# converting_exel_files_to_list_PM_for_OG()
# converting_exel_files_to_list_for_sog()
# converting_exel_files_to_list_for_og()
# converting_exel_files_to_list_for_PAT()
# converting_exel_files_to_list_AK_for_PAT()


class App(tk.Tk):
    #  const
    OLDER_OFFICERS_OG = 'older_og'
    OFFICERS_OG = 'officer_og'
    GROUP_PAT = 'pat'
    PM_FOR_OG = 'pm_for_group_og'
    AK_FOR_PAT = 'ak_for_pat'

    SHORTED_OG = 'shorted_og'
    PM_FULL_OG = 'pm_full_og'


    dict_for_render = {}
    dict_for_render[OFFICERS_OG] = ['', '', '', '', '', '']
    dict_for_render[SHORTED_OG] = ['', '']
    dict_for_render[PM_FOR_OG] = ['', '']
    dict_for_render[PM_FULL_OG] = ['', '', '', '', '', '', '', '']

    def __init__(self):
        super().__init__()
        self.title('GUI for work')
        self.geometry('900x800+250+0')
        self.iconbitmap('arm.ico')
        self.resizable(False, False)

        self.elements_lists = {}
        self.selected_ogs = []

        self.label_month = ttk.Label(text="Выберете месяц:").grid()
        self.combobox_month = ttk.Combobox(values=['январь', 'февраль', 'март', 'апреля', 'май', 'июнь', 'июль',
                                                   'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'], width=40)
        self.combobox_month.grid()
        self.combobox_month.bind('<<ComboboxSelected>>', self.add_entry_date)

        self.label_days = ttk.Label(text="Выберете дату:").grid()
        self.combobox_days = ttk.Combobox(values=[i for i in range(1, 32)], width=40)
        self.combobox_days.grid(padx=10, pady=1)
        self.combobox_days.bind('<<ComboboxSelected>>', self.add_entry_date)
        self.combobox_day = ttk.Combobox(values=[i for i in range(1, 32)], width=40)
        self.combobox_day.grid(pady=1)
        self.combobox_day.bind('<<ComboboxSelected>>', self.add_entry_date)

        # OG SHORTED
        self.label_shorted_og = ttk.Label(text="Выберете ОГ сокращенного состава:").grid()
        self.combobox_sog = ttk.Combobox(values=xls_loader.data[self.OLDER_OFFICERS_OG], width=40,
                                         name='cbox_shorted_og_1')
        self.combobox_sog.grid(pady=1)
        self.combobox_sog.bind('<<ComboboxSelected>>', self.handle_shorted_og_event)
        self.combobox_og = ttk.Combobox(values=xls_loader.data[self.OFFICERS_OG], width=40, name='cbox_shorted_og_2')
        self.combobox_og.grid(pady=1)
        self.combobox_og.bind('<<ComboboxSelected>>', self.handle_shorted_og_event)

        # OG FULL
        self.label_full_squad = ttk.Label(text="Выберете ОГ полного состава:").grid()
        self.combobox_full_sog = ttk.Combobox(values=xls_loader.data[self.OLDER_OFFICERS_OG], width=40,
                                              name='cbox_sog_full_og')
        self.combobox_full_sog.grid(pady=1)
        self.combobox_full_sog.bind('<<ComboboxSelected>>', self.handle_og_event)
        for item in range(1, 7):
            self.com = ttk.Combobox(values=xls_loader.data[self.OFFICERS_OG], width=40, name=f'cbox_fog_{item}')
            self.com.grid(pady=1)
            self.com.bind('<<ComboboxSelected>>', self.handle_og_event)

        # PM
        self.label_shorted_PM = ttk.Label(text="Выберете PM для ОГ сокращенного состава:").grid()
        self.combobox_pm_for_shorted_sog = ttk.Combobox(values=xls_loader.data[self.PM_FOR_OG], width=40,
                                                        name='cbox_pm_for_shorted_og_1')
        self.combobox_pm_for_shorted_sog.grid(pady=2)
        self.combobox_pm_for_shorted_sog.bind('<<ComboboxSelected>>', self.handle_pm_for_shorted_og_event)

        self.combobox_pm_for_shorted_og = ttk.Combobox(values=xls_loader.data[self.PM_FOR_OG], width=40,
                                                       name='cbox_pm_for_shorted_og_2')
        self.combobox_pm_for_shorted_og.grid(pady=1)
        self.combobox_pm_for_shorted_og.bind('<<ComboboxSelected>>', self.handle_pm_for_shorted_og_event)

        self.label_full_PM = ttk.Label(text="Выберете PM для ОГ полного состава:").grid()
        for item in range(1, 8):
            self.com = ttk.Combobox(values=xls_loader.data[self.PM_FOR_OG], width=40, name=f'cbox_pm_fog_{item}')
            self.com.grid(pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_full_og_pm)

        # PAT
        self.label_PAT = ttk.Label(text="Выберете PAT состава:").grid(row=0, column=2)
        self.label_control_group = ttk.Label(text="Control group").grid(row=1, column=2)
        for item in range(2, 4):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2, padx=40, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_ranger_group = ttk.Label(text="ranger group").grid(row=4, column=2)
        for item in range(5, 7):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_defense_group = ttk.Label(text="defense group").grid(row=7, column=2)
        for item in range(8, 16):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_fire_group = ttk.Label(text="fire group").grid(row=16, column=2)
        for item in range(17, 20):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_reserve_group = ttk.Label(text="reserve group").grid(row=20, column=2)
        for item in range(21, 25):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_sappers_group = ttk.Label(text="sappers group").grid(row=25, column=2)
        for item in range(26, 28):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_rhb_group = ttk.Label(text="rhb group").grid(row=28, column=2)
        for item in range(29, 31):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        self.label_med_group = ttk.Label(text="med group").grid(row=31, column=2)
        for item in range(32, 34):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], width=40)
            self.com.grid(row=item, column=2, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat)

        # AK FOR PAT
        self.label_ak_PAT = ttk.Label(text="Выберете ak PAT :").grid(row=0, column=3)
        self.label_ak_control_group = ttk.Label(text="Control group").grid(row=1, column=3)
        for item in range(2, 4):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_ranger_group = ttk.Label(text="ranger group").grid(row=4, column=3)
        for item in range(5, 7):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_defense_group = ttk.Label(text="defense group").grid(row=7, column=3)
        for item in range(8, 16):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_fire_group = ttk.Label(text="fire group").grid(row=16, column=3)
        for item in range(17, 20):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_reserve_group = ttk.Label(text="reserve group").grid(row=20, column=3)
        for item in range(21, 25):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_sappers_group = ttk.Label(text="sappers group").grid(row=25, column=3)
        for item in range(26, 28):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_rhb_group = ttk.Label(text="rhb group").grid(row=28, column=3)
        for item in range(29, 31):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.label_ak_med_group = ttk.Label(text="med group").grid(row=31, column=3)
        for item in range(32, 34):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], width=40)
            self.com.grid(row=item, column=3)
            self.com.bind('<<ComboboxSelected>>', self.add_entry_pat_ak)

        self.btn = tk.Button(text='Сгенерировать документ', width=36, heigh=6, bg='green')
        self.btn.place(relx=0.01, rely=0.8)
        self.btn.bind('<Button-1>', self.render_template)

    def add_entry_date(self, event):
        element = event.widget.get()
        if event:
            if element == 'август' or element == 'март':
                element += 'a'
                list_for_render_date_and_month.append(element)
                print(list_for_render_date_and_month)
            elif element == 'май':
                element = 'мая'
                list_for_render_date_and_month.append(element)
                print(list_for_render_date_and_month)
            else:
                element = element.replace(element[-1], 'я')
                list_for_render_date_and_month.append(element)
                print(list_for_render_date_and_month)

    def handle_shorted_og_event(self, event):
        element = event.widget.get()
        if event:
            print(f'выбрано: {event.widget.get()}')
            list_for_render_shorted_og.append(event.widget.get())
            index_name = int(str(event.widget).split("_")[-1])
            self.dict_for_render[self.SHORTED_OG][index_name - 1] = element

    # def add_entry_full_og(self, event):
    #     element = event.widget.get()
    #     if event:
    #         if element in list_of_og:
    #             print(f"Element '{element}' in list_of_og")
    #             list_for_render_full_og.append(element)
    #             list_of_og.remove(element)
    #             print(f"Element '{element}' add in list_for_render_full_og")
    #         else:
    #             for element_render in list_for_render_full_og:
    #                 if element_render == element:
    #                     self._show_info()

    def handle_og_event(self, event):
        element = event.widget.get()
        if str(event.widget).split(".")[-1] == 'cbox_sog_full_og':
            self.dict_for_render[self.OLDER_OFFICERS_OG] = [element]
        elif str(event.widget).split(".")[-1].startswith('cbox_fog'):
            index_name = int(str(event.widget).split("_")[-1])
            print(index_name)
            self.dict_for_render[self.OFFICERS_OG][index_name - 1] = element
            print(self.dict_for_render[self.OFFICERS_OG])
        # TODO доработать проверку на наличие повторяемых (выбранных) значений

    def handle_pm_for_shorted_og_event(self, event):
        element = event.widget.get()
        if event:
            if str(event.widget).split(".")[-1].startswith('cbox_pm_for_shorted_og'):
                list_for_render_shorted_og_pm.append(element)
                index_name = int(str(event.widget).split("_")[-1])
                print(index_name)
                self.dict_for_render[self.PM_FOR_OG][index_name - 1] = element
                print(self.dict_for_render[self.PM_FOR_OG])
        # TODO доработать проверку на наличие повторяемых (выбранных) значений

    def add_entry_full_og_pm(self, event):
        element = event.widget.get()
        if event:
            if str(event.widget).split(".")[-1].startswith('cbox_pm_fog'):
                index_name = int(str(event.widget).split("_")[-1])
                print(index_name)
                self.dict_for_render[self.PM_FULL_OG][index_name - 1] = element
                print(self.dict_for_render[self.PM_FULL_OG])


            # if element in list_PM:
            #     print(f"Element '{element}' in list_PM")
            #     list_for_render_full_og_pm.append(element)
            #     print(f"Element '{element}' ADD in list_for_render_full_og")
            #     list_PM.remove(element)
            #     print(f"Element '{element}' REMOVE in list_PM")
            # else:
            #     for element_render in list_for_render_full_og_pm:
            #         if element_render == element:
            #             self._show_info()

    def add_entry_pat(self, event):
        element = event.widget.get()
        if event:
            if element in list_of_pat:
                print(f"Element '{element}' in list_of_pat")
                list_for_render_pat.append(element)
                print(f"Element '{element}' ADD in list_for_render_pat")
                list_of_pat.remove(element)
                print(f"Element '{element}' REMOVE in list_of_pat")
            else:
                for element_render in list_for_render_pat:
                    if element_render == element:
                        self._show_info()

    def add_entry_pat_ak(self, event):
        element = event.widget.get()
        if event:
            if element in list_AK:
                print(f"Element '{element}' in list_AK")
                list_for_render_pat_ak.append(element)
                print(f"Element '{element}' ADD in list_for_render_pat_ak")
                list_AK.remove(element)
                print(f"Element '{element}' REMOVE in list_AK")
            else:
                for element_render in list_for_render_pat_ak:
                    if element_render == element:
                        self._show_info()

    def _show_info(self):
        msg = "Запись уже существует! Поменяйте свой выбор."
        mb.showinfo("Информация", msg)

    def render_template(self, event):
        # dict = {index: value for index, value in enumerate(lst}
        context_list = {
            'month': list_for_render_date_and_month[0],
            'number1': list_for_render_date_and_month[1],
            'number2': list_for_render_date_and_month[1],
            'og0': list_for_render_shorted_og[0],
            'og1': list_for_render_shorted_og[1],
            'og2': list_for_render_full_og[0],
            'og3': list_for_render_full_og[1],
            'og4': list_for_render_full_og[2],
            'og5': list_for_render_full_og[3],
            'og6': list_for_render_full_og[4],
            'og7': list_for_render_full_og[5],
            'og8': list_for_render_full_og[6],
            # 'pat0': list_for_render_pat[0],
            # 'pat1': list_for_render_pat[1],
            # 'pat2': list_for_render_pat[2],
            # 'pat3': list_for_render_pat[3],
            # 'pat4': list_for_render_pat[4],
            # 'pat5': list_for_render_pat[5],
            # 'pat6': list_for_render_pat[6],
            # 'pat7': list_for_render_pat[7],
            # 'pat8': list_for_render_pat[8],
            # 'pat9': list_for_render_pat[9],
            # 'pat10': list_for_render_pat[10],
            # 'pat11': list_for_render_pat[11],
            # 'pat12': list_for_render_pat[12],
            # 'pat13': list_for_render_pat[13],
            # 'pat14': list_for_render_pat[14],
            # 'pat15': list_for_render_pat[15],
            # 'pat16': list_for_render_pat[16],
            # 'pat17': list_for_render_pat[17],
            # 'pat18': list_for_render_pat[18],
            # 'pat19': list_for_render_pat[19],
            # 'pat20': list_for_render_pat[20],
            # 'pat21': list_for_render_pat[21],
            # 'pat22': list_for_render_pat[22],
            # 'pat23': list_for_render_pat[23],
            # 'pat24': list_for_render_pat[24],
            # 'pm0': list_for_render_shorted_og_pm[0],
            # 'pm1': list_for_render_shorted_og_pm[1],
            # 'pm2': list_for_render_full_og_pm[0],
            # 'pm3': list_for_render_full_og_pm[1],
            # 'pm4': list_for_render_full_og_pm[2],
            # 'pm5': list_for_render_full_og_pm[3],
            # 'pm6': list_for_render_full_og_pm[4],
            # 'pm7': list_for_render_full_og_pm[5],
            # 'pm8': list_for_render_full_og_pm[6],
            # 'ak0': list_for_render_pat_ak[0],
            # 'ak1': list_for_render_pat_ak[1],
            # 'ak2': list_for_render_pat_ak[2],
            # 'ak3': list_for_render_pat_ak[3],
            # 'ak4': list_for_render_pat_ak[4],
            # 'ak5': list_for_render_pat_ak[5],
            # 'ak6': list_for_render_pat_ak[6],
            # 'ak7': list_for_render_pat_ak[7],
            # 'ak8': list_for_render_pat_ak[8],
            # 'ak9': list_for_render_pat_ak[9],
            # 'ak10': list_for_render_pat_ak[10],
            # 'ak11': list_for_render_pat_ak[11],
            # 'ak12': list_for_render_pat_ak[12],
            # 'ak13': list_for_render_pat_ak[13],
            # 'ak14': list_for_render_pat_ak[14],
            # 'ak15': list_for_render_pat_ak[15],
            # 'ak16': list_for_render_pat_ak[16],
            # 'ak17': list_for_render_pat_ak[17],
            # 'ak18': list_for_render_pat_ak[18],
            # 'ak19': list_for_render_pat_ak[19],
            # 'ak20': list_for_render_pat_ak[20],
            # 'ak21': list_for_render_pat_ak[21],
            # 'ak22': list_for_render_pat_ak[22],
            # 'ak23': list_for_render_pat_ak[23],
            # 'ak24': list_for_render_pat_ak[24],
        }
        print('exelent!')
        doc.render(context_list)
        doc.save('probe3.docx')


app = App()
app.mainloop()
