import tkinter as tk
from tkinter import ttk


class WindowWeek(tk.Toplevel):

    def __init__(self, master=None):
        super().__init__(master=master)
        self.title("New Window")
        self.geometry("200x200")
        label = tk.Label(self, text="This is a new Window2")
        label.pack()