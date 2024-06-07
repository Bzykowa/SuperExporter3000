import tkinter as tk


class MainMenu(tk.Frame):
    """The main menu of the application."""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.setup()

    def setup(self):
        """Shape the menu component"""
        self.btn_delegation = tk.Button(self, text="Delegacje")
