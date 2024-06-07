import tkinter as tk
from gui.main_menu import MainMenu


class MainWindow(tk.Frame):
    """The main widget of the application"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.setup()

    def setup(self):
        """Shape the main window"""
        self.master.title("SuperImporter3000")
        self.master.rowconfigure(0, weight=1, minsize=800)
        self.master.columnconfigure(1, weight=1, minsize=800)

        self.main_menu = MainMenu(self)
