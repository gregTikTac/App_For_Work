import tkinter as tk
from tkinter.constants import *
from tkinter import ttk
from docxtpl import DocxTemplate
from XlsLoader import XlsLoader
from WindowDay import WindowDay
from WindowWeek import WindowWeek


#  шаблон документа
doc = DocxTemplate('probe3.docx')

# читаем файл
xls_loader = XlsLoader('word_automation.xlsm')
xls_loader.load()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('GUI for work')
        self.geometry('610x490+450+150')
        self.iconbitmap('arm.ico')
        self.resizable(False, False)

        self.gif_image = tk.PhotoImage(file='s_640.gif')
        self.label_panel = ttk.Label(image=self.gif_image)
        self.label_panel.place(x=50, y=10, relwidth=1, relheight=1)

        # _____________________________________________buttons__________________________________________________
        self.btn_day = ttk.Button(text='Ежедневный')
        self.btn_day.place(width=150, height=50, x=100, y=200)
        self.btn_day.bind("<Button-1>", lambda x: WindowDay(self))

        self.btn_week = ttk.Button(text='Еженедельный')
        self.btn_week.place(width=150, height=50, x=360, y=200)
        self.btn_week.bind("<Button-1>", lambda e: WindowWeek(self))

    def run_app(self):
        self.mainloop()


if __name__ == '__main__':
    app = App()
    app.run_app()

