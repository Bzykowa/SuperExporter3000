import tkinter as tk

from gui.delegations_ui import DelegationsUI
from gui.invoices_ui import InvoicesUI


class MainMenu(tk.Frame):
    """The main menu of the application."""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.controller = controller
        self.rowconfigure((0, 1), weight=1)
        self.columnconfigure(0, weight=1)
        self.setup()

    def setup(self):
        """Shape the menu component"""
        # Components
        self.config(relief=tk.RAISED, bd=2, bg="azure")
        self.btn_delegation = tk.Button(
            self, text="Delegacje", command=self.open_del_ui)
        self.btn_invoice = tk.Button(
            self, text="Faktury", command=self.open_inv_ui)
        # Placement
        self.btn_delegation.grid(
            row=0, column=0, sticky="ew", padx=15, pady=5)
        self.btn_invoice.grid(row=1, column=0, sticky="ew", padx=15, pady=5)

    def open_del_ui(self):
        """Show the frame containing delegation import components."""
        self.controller.show_page(DelegationsUI)

    def open_inv_ui(self):
        """Show the frame containing invoices import components."""
        self.controller.show_page(InvoicesUI)
