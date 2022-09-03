import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as mb


class WindowDay(tk.Toplevel):
    """создает дочернее окно ежедневного документа"""
    def __init__(self, master=None):
        super().__init__(master=master)
        self.title("DOC on Day")
        self.geometry('1000x780+250+0')
        self.iconbitmap('arm.ico')
        self.resizable(False, False)

        # _____________________buttons and labels_____________________________________

        #  WM_DELETE_WINDOW вызывается, когда окно верхнего уровня должно уже закрываться, и по умолчанию Tk уничтожает
        #  окно, для которого оно было получено
        self.protocol("WM_DELETE_WINDOW", self.confirm_delete)
        self.grab_focus()  # вызов

    def grab_focus(self):
        """Задерживает роительское окно, не позволяя выполнять другие действия"""
        self.grab_set()
        self.focus_set()
        self.wait_window()

    def confirm_delete(self):
        """При закрытии окна, спрашивает стоит ли закрывать"""
        message = "Вы уверены, что хотите закрыть это окно?"
        if mb.askyesno(message=message, parent=self):
            self.destroy()
