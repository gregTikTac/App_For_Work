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
# list_of_sog = []
# list_of_og = []
# list_of_pat = []
# list_PM = []
# list_AK = []
list_for_render_date_and_month = []


# list_for_render_shorted_og = []
# list_for_render_full_og = []
# list_for_render_shorted_og_pm = []
# list_for_render_full_og_pm = []
# list_for_render_pat = []
# list_for_render_pat_ak = []


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

    dict_for_render = {
        OFFICERS_OG: ['', '', '', '', '', ''],
        SHORTED_OG: ['', ''],
        PM_FOR_OG: ['', ''],
        PM_FULL_OG: ['', '', '', '', '', '', '', ''],
        GROUP_PAT: ['' for i in range(1, 26)],
        AK_FOR_PAT: ['' for i in range(1, 26)]
    }

    def __init__(self):
        super().__init__()
        self.title('GUI for work')
        self.geometry('1000x800+250+0')
        self.iconbitmap('arm.ico')
        self.resizable(False, False)

        self.elements_lists = {}
        self.selected_ogs = []

        self.label_month = ttk.Label(text="Выберете месяц:").grid()
        self.combobox_month = ttk.Combobox(values=['январь', 'февраль', 'март', 'апреля', 'май', 'июнь', 'июль',
                                                   'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'], width=40)
        self.combobox_month.grid()
        self.combobox_month.bind('<<ComboboxSelected>>', self.add_entry_month)

        self.label_days = ttk.Label(text="Выберете дату:").grid()
        self.combobox_days = ttk.Combobox(values=[i for i in range(1, 32)], width=40)
        self.combobox_days.grid(padx=10, pady=1)
        self.combobox_days.bind('<<ComboboxSelected>>', self.add_date)
        self.combobox_day = ttk.Combobox(values=[i for i in range(1, 32)], width=40)
        self.combobox_day.grid(pady=1)
        self.combobox_day.bind('<<ComboboxSelected>>', self.add_date)

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
            self.com_fog = ttk.Combobox(values=xls_loader.data[self.OFFICERS_OG], width=40, name=f'cbox_fog_{item}')
            self.com_fog.grid(pady=1)
            self.com_fog.bind('<<ComboboxSelected>>', self.handle_og_event)


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
            self.com.bind('<<ComboboxSelected>>', self.handle_pm_for_full_og_event)

        # PAT
        self.label_PAT = ttk.Label(text="Выберете PAT состава:").place(x=425, y=1)
        self.label_contol_group = ttk.Label(text="Control:").place(x=300, y=20)
        self.label_ranger_group = ttk.Label(text="Rangers:").place(x=300, y=65)
        self.label_block_group = ttk.Label(text="Block:").place(x=310, y=110)
        self.label_fire_group = ttk.Label(text="Fire:").place(x=310, y=295)
        self.label_reserve_group = ttk.Label(text="Reserve:").place(x=300, y=365)
        self.label_sapper_group = ttk.Label(text="Sappers:").place(x=300, y=460)
        self.label_rhb_group = ttk.Label(text="Rhb:").place(x=310, y=505)
        self.label_med_group = ttk.Label(text="Med:").place(x=310, y=550)

        for item in range(1, 26):
            self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
            self.com.grid(row=item, column=3, padx=70, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)

        # PAT AK
        self.label_PAT = ttk.Label(text="Выберете PAT состава:").place(x=765, y=1)
        self.label_contol_group = ttk.Label(text="Control:").place(x=640, y=20)
        self.label_ranger_group = ttk.Label(text="Rangers:").place(x=640, y=65)
        self.label_block_group = ttk.Label(text="Block:").place(x=650, y=110)
        self.label_fire_group = ttk.Label(text="Fire:").place(x=650, y=295)
        self.label_reserve_group = ttk.Label(text="Reserve:").place(x=640, y=365)
        self.label_sapper_group = ttk.Label(text="Sappers:").place(x=640, y=460)
        self.label_rhb_group = ttk.Label(text="Rhb:").place(x=650, y=505)
        self.label_med_group = ttk.Label(text="Med:").place(x=650, y=550)

        for item in range(1, 26):
            self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
            self.com.grid(row=item, column=4, padx=10, pady=1)
            self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # # self.label_ranger_group = ttk.Label(text="ranger group").place(x=415, y=85)
        # for item in range(4, 6):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)
        #
        # # self.label_defense_group = ttk.Label(text="defense group").place(x=415, y=155)
        # for item in range(7, 15):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)
        #
        # # self.label_fire_group = ttk.Label(text="fire group").place(x=425, y=360)
        # for item in range(16, 19):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2, pady=1)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)
        #
        # # self.label_reserve_group = ttk.Label(text="reserve group").place(x=415, y=450)
        # for item in range(20, 24):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2, pady=1)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)
        #
        # self.label_sappers_group = ttk.Label(text="sappers group").place(x=415, y=565)
        # for item in range(25, 27):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2, pady=1)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)
        #
        # self.label_rhb_group = ttk.Label(text="rhb group").place(x=425, y=635)
        # for item in range(28, 30):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2, pady=1)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)
        #
        # self.label_med_group = ttk.Label(text="med group").place(x=425, y=697)
        # for item in range(31, 33):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.GROUP_PAT], name=f'cbox_pat_{item}', width=40)
        #     self.com.grid(row=item+1, column=2, pady=1)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_pat_event)

        # # AK FOR PAT
        # self.label_ak_PAT = ttk.Label(text="Выберете ak PAT :").grid(row=0, column=3)
        # self.label_ak_control_group = ttk.Label(text="Control group").grid(row=1, column=3)
        # for item in range(2, 4):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_ranger_group = ttk.Label(text="ranger group").grid(row=4, column=3)
        # for item in range(5, 7):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_defense_group = ttk.Label(text="defense group").grid(row=7, column=3)
        # for item in range(8, 16):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_fire_group = ttk.Label(text="fire group").grid(row=16, column=3)
        # for item in range(17, 20):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_reserve_group = ttk.Label(text="reserve group").grid(row=20, column=3)
        # for item in range(21, 25):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_sappers_group = ttk.Label(text="sappers group").grid(row=25, column=3)
        # for item in range(26, 28):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_rhb_group = ttk.Label(text="rhb group").grid(row=28, column=3)
        # for item in range(29, 31):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)
        #
        # self.label_ak_med_group = ttk.Label(text="med group").grid(row=31, column=3)
        # for item in range(32, 34):
        #     self.com = ttk.Combobox(values=xls_loader.data[self.AK_FOR_PAT], name=f'cbox_ak_pat_{item}', width=40)
        #     self.com.grid(row=item, column=3)
        #     self.com.bind('<<ComboboxSelected>>', self.handle_ak_for_pat_event)

        self.btn = tk.Button(text='Сгенерировать документ', width=36, heigh=6, bg='green')
        self.btn.place(relx=0.01, rely=0.8)
        self.btn.bind('<Button-1>', self.render_template)

    def add_date(self, event):
        element = event.widget.get()
        list_for_render_date_and_month.append(element)
        print(list_for_render_date_and_month)

    def add_entry_month(self, event):
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
        new_lst = []
        if event:
            # print(f'выбрано: {event.widget.get()}')
            # list_for_render_shorted_og.append(event.widget.get())
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
            print(self.dict_for_render[self.OLDER_OFFICERS_OG])
        elif str(event.widget).split(".")[-1].startswith('cbox_fog'):
            index_name = int(str(event.widget).split("_")[-1])
            if element not in self.dict_for_render[self.OFFICERS_OG]:
            # for option in xls_loader.data[self.OFFICERS_OG]:
            #     if option != element:
            #         self.new_data.append(option)
            # xls_loader.data[self.OFFICERS_OG] = self.new_data
            # print(xls_loader.data[self.OFFICERS_OG])
                self.dict_for_render[self.OFFICERS_OG][index_name - 1] = element
                print(self.dict_for_render[self.OFFICERS_OG])
            else:
                self._show_info()


        # TODO доработать проверку на наличие повторяемых (выбранных) значений(обновление values)

    def handle_pm_for_shorted_og_event(self, event):
        element = event.widget.get()
        if event:
            if str(event.widget).split(".")[-1].startswith('cbox_pm_for_shorted_og'):
                index_name = int(str(event.widget).split("_")[-1])
                print(index_name)
                self.dict_for_render[self.PM_FOR_OG][index_name - 1] = element
                print(self.dict_for_render[self.PM_FOR_OG])
        # TODO доработать проверку на наличие повторяемых (выбранных) значений

    def handle_pm_for_full_og_event(self, event):
        element = event.widget.get()
        if event:
            if str(event.widget).split(".")[-1].startswith('cbox_pm_fog'):
                index_name = int(str(event.widget).split("_")[-1])
                if element not in self.dict_for_render[self.PM_FULL_OG]:
                    self.dict_for_render[self.PM_FULL_OG][index_name - 1] = element
                    print(self.dict_for_render[self.PM_FULL_OG])
                else:
                    self._show_info()

        # TODO доработать проверку на наличие повторяемых (выбранных) значений

    def handle_pat_event(self, event):
        element = event.widget.get()
        if event:
            if str(event.widget).split(".")[-1].startswith('cbox_pat'):
                index_name = int(str(event.widget).split("_")[-1])
                if element not in self.dict_for_render[self.GROUP_PAT]:
                    self.dict_for_render[self.GROUP_PAT][index_name - 1] = element
                    print(self.dict_for_render[self.GROUP_PAT])
                else:
                    self._show_info()

    def handle_ak_for_pat_event(self, event):
        element = event.widget.get()
        if event:
            if str(event.widget).split(".")[-1].startswith('cbox_ak_pat'):
                index_name = int(str(event.widget).split("_")[-1])
                print(index_name)
                if element not in self.dict_for_render[self.AK_FOR_PAT]:
                    self.dict_for_render[self.AK_FOR_PAT][index_name - 1] = element
                    print(self.dict_for_render[self.AK_FOR_PAT])
                else:
                    self._show_info()

    # def add_entry_full_og_pm(self, event):
    #     element = event.widget.get()
    #     if event:
    #         if str(event.widget).split(".")[-1].startswith('cbox_pm_fog'):
    #             index_name = int(str(event.widget).split("_")[-1])
    #             print(index_name)
    #             self.dict_for_render[self.PM_FULL_OG][index_name - 1] = element
    #             print(self.dict_for_render[self.PM_FULL_OG])

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

    # def add_entry_pat(self, event):
    #     element = event.widget.get()
    #     if event:
    #         if element in list_of_pat:
    #             print(f"Element '{element}' in list_of_pat")
    #             list_for_render_pat.append(element)
    #             print(f"Element '{element}' ADD in list_for_render_pat")
    #             list_of_pat.remove(element)
    #             print(f"Element '{element}' REMOVE in list_of_pat")
    #         else:
    #             for element_render in list_for_render_pat:
    #                 if element_render == element:
    #                     self._show_info()

    # def add_entry_pat_ak(self, event):
    #     element = event.widget.get()
    #     if event:
    #         if element in list_AK:
    #             print(f"Element '{element}' in list_AK")
    #             list_for_render_pat_ak.append(element)
    #             print(f"Element '{element}' ADD in list_for_render_pat_ak")
    #             list_AK.remove(element)
    #             print(f"Element '{element}' REMOVE in list_AK")
    #         else:
    #             for element_render in list_for_render_pat_ak:
    #                 if element_render == element:
    #                     self._show_info()

    def _show_info(self):
        msg = "Запись уже существует! Поменяйте свой выбор."
        mb.showinfo("Информация", msg)

    def render_template(self, event):
        self.dict_for_render = {
            'month': list_for_render_date_and_month[0],
            'number1': list_for_render_date_and_month[1],
            'number2': list_for_render_date_and_month[2],
            'shorted_og': self.dict_for_render[self.SHORTED_OG][0],
            'shorted_og2': self.dict_for_render[self.SHORTED_OG][1],
            'full_sog': self.dict_for_render[self.OLDER_OFFICERS_OG][0],
            'full_og1': self.dict_for_render[self.OFFICERS_OG][0],
            'full_og2': self.dict_for_render[self.OFFICERS_OG][1],
            'full_og3': self.dict_for_render[self.OFFICERS_OG][2],
            'full_og4': self.dict_for_render[self.OFFICERS_OG][3],
            'full_og5': self.dict_for_render[self.OFFICERS_OG][4],
            'full_og6': self.dict_for_render[self.OFFICERS_OG][5],
            'pat0': self.dict_for_render[self.GROUP_PAT][0],
            'pat1': self.dict_for_render[self.GROUP_PAT][1],
            'pat2': self.dict_for_render[self.GROUP_PAT][2],
            'pat3': self.dict_for_render[self.GROUP_PAT][3],
            'pat4': self.dict_for_render[self.GROUP_PAT][4],
            'pat5': self.dict_for_render[self.GROUP_PAT][5],
            'pat6': self.dict_for_render[self.GROUP_PAT][6],
            'pat7': self.dict_for_render[self.GROUP_PAT][7],
            'pat8': self.dict_for_render[self.GROUP_PAT][8],
            'pat9': self.dict_for_render[self.GROUP_PAT][9],
            'pat10': self.dict_for_render[self.GROUP_PAT][10],
            'pat11': self.dict_for_render[self.GROUP_PAT][11],
            'pat12': self.dict_for_render[self.GROUP_PAT][12],
            'pat13': self.dict_for_render[self.GROUP_PAT][13],
            'pat14': self.dict_for_render[self.GROUP_PAT][14],
            'pat15': self.dict_for_render[self.GROUP_PAT][15],
            'pat16': self.dict_for_render[self.GROUP_PAT][16],
            'pat17': self.dict_for_render[self.GROUP_PAT][17],
            'pat18': self.dict_for_render[self.GROUP_PAT][18],
            'pat19': self.dict_for_render[self.GROUP_PAT][19],
            'pat20': self.dict_for_render[self.GROUP_PAT][20],
            'pat21': self.dict_for_render[self.GROUP_PAT][21],
            'pat22': self.dict_for_render[self.GROUP_PAT][22],
            'pat23': self.dict_for_render[self.GROUP_PAT][23],
            'pat24': self.dict_for_render[self.GROUP_PAT][24],
            'pm0': self.dict_for_render[self.PM_FOR_OG][0],
            'pm1': self.dict_for_render[self.PM_FOR_OG][1],
            'pm2': self.dict_for_render[self.PM_FULL_OG][0],
            'pm3': self.dict_for_render[self.PM_FULL_OG][1],
            'pm4': self.dict_for_render[self.PM_FULL_OG][2],
            'pm5': self.dict_for_render[self.PM_FULL_OG][3],
            'pm6': self.dict_for_render[self.PM_FULL_OG][4],
            'pm7': self.dict_for_render[self.PM_FULL_OG][5],
            'pm8': self.dict_for_render[self.PM_FULL_OG][6],
            'ak0': self.dict_for_render[self.AK_FOR_PAT][0],
            'ak1': self.dict_for_render[self.AK_FOR_PAT][1],
            'ak2': self.dict_for_render[self.AK_FOR_PAT][2],
            'ak3': self.dict_for_render[self.AK_FOR_PAT][3],
            'ak4': self.dict_for_render[self.AK_FOR_PAT][4],
            'ak5': self.dict_for_render[self.AK_FOR_PAT][5],
            'ak6': self.dict_for_render[self.AK_FOR_PAT][6],
            'ak7': self.dict_for_render[self.AK_FOR_PAT][7],
            'ak8': self.dict_for_render[self.AK_FOR_PAT][8],
            'ak9': self.dict_for_render[self.AK_FOR_PAT][9],
            'ak10': self.dict_for_render[self.AK_FOR_PAT][10],
            'ak11': self.dict_for_render[self.AK_FOR_PAT][11],
            'ak12': self.dict_for_render[self.AK_FOR_PAT][12],
            'ak13': self.dict_for_render[self.AK_FOR_PAT][13],
            'ak14': self.dict_for_render[self.AK_FOR_PAT][14],
            'ak15': self.dict_for_render[self.AK_FOR_PAT][15],
            'ak16': self.dict_for_render[self.AK_FOR_PAT][16],
            'ak17': self.dict_for_render[self.AK_FOR_PAT][17],
            'ak18': self.dict_for_render[self.AK_FOR_PAT][18],
            'ak19': self.dict_for_render[self.AK_FOR_PAT][19],
            'ak20': self.dict_for_render[self.AK_FOR_PAT][20],
            'ak21': self.dict_for_render[self.AK_FOR_PAT][21],
            'ak22': self.dict_for_render[self.AK_FOR_PAT][22],
            'ak23': self.dict_for_render[self.AK_FOR_PAT][23],
            'ak24': self.dict_for_render[self.AK_FOR_PAT][24],
        }
        # print('exelent!')
        doc.render(self.dict_for_render)
        doc.save('probe3.docx')
        print(self.dict_for_render)


app = App()
app.mainloop()
