# import tkinter as tk
# from tkinter import ttk
#
#
# class App(tk.Tk):
#     def __init__(self):
#         super().__init__()
#
#         self.title('GUI for work')
#         self.geometry('980x700+250+0')
#         self.iconbitmap('arm.ico')
#         self.resizable(False, False)


import tkinter as tk
from tkinter.constants import *
from tkinter import ttk


class App(tk.Tk):
    BKGR_IMAGE_PATH = 's.gif'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title('GUI for work')
        self.geometry('1000x780+250+0')
        self.iconbitmap('arm.ico')
        self.resizable(False, False)

        main_frame = tk.Frame(self)
        main_frame.pack(side='top', fill='both', expand='True')

        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        self.bkgr_image = tk.PhotoImage(file=self.BKGR_IMAGE_PATH)

        self.frames = {}
        for F in (PageOne, PageTwo):
            frame = F(main_frame, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky='nsew')

        self.show_frame(PageOne)

    def show_frame(self, container):
        frame = self.frames[container]
        frame.tkraise()


class BasePage(tk.Frame):
    """ Расположение звезды на фоне """

    def __init__(self, parent, controller):
        super().__init__(parent)

        label_bkgr = tk.Label(self, image=controller.bkgr_image)
        label_bkgr.place(relx=0.5, rely=0.5, anchor=CENTER)  # Center label w/image.


class PageOne(BasePage):
    """ Страница для создания суточного жкумента"""

    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        button_for_jump_to_week = ttk.Button(self, text="Переход к еженедельному",
                                             command=lambda: controller.show_frame(PageTwo))
        button_for_jump_to_week.pack(side=LEFT, anchor='nw')

        label1 = ttk.Label(self, text='Ежедневный', font=("Times New Roman", 16))
        label1.pack(side=LEFT, anchor='nw', padx=10)


class PageTwo(BasePage):

    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        label2 = ttk.Label(self, text='Page 2', font=("Helvetica", 20))
        label2.pack(padx=10, pady=10)

        button4 = ttk.Button(self, text="Page 1",
                             command=lambda: controller.show_frame(PageOne))
        button4.pack()


app = App()
app.mainloop()
