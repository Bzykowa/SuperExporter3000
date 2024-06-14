import tkinter as tk
from gui.delegations_ui import DelegationsUI
from gui.invoices_ui import InvoicesUI
from gui.main_menu import MainMenu


class MainWindow(tk.Frame):
    """The main widget of the application"""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.pack()
        self.setup()

    def setup(self):
        """Shape the main window"""
        self.master.title("SuperImporter3000")
        self.master.rowconfigure(0, weight=1, minsize=800)
        self.master.columnconfigure(1, weight=1, minsize=800)
        self.frames = {}

        # Add main menu
        self.frames[MainMenu] = MainMenu(self, self)
        self.frames[MainMenu].grid(row=0, column=0, sticky="ns")
        # Add side frames
        self.frames[DelegationsUI] = DelegationsUI(self)
        self.frames[DelegationsUI].grid(row=0, column=1, sticky="nsew")
        self.frames[InvoicesUI] = InvoicesUI(self)
        self.frames[InvoicesUI].grid(row=0, column=1, sticky="nsew")
        self.show_page(DelegationsUI)

    def get_page(self, page_class):
        return self.frames[page_class]

    def show_page(self, page_class):
        self.frames[page_class].tkraise()
