import tkinter as tk
import pandas as pd
import config.utils as cfg
import pathlib
from tkinter.filedialog import askdirectory
from tkcalendar import DateEntry

from xml_parsing.invoices import Invoices


class InvoicesUI(tk.Frame):
    """The part of the application responsible for setting up invoices
    import."""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.controller = controller
        self.rowconfigure((0, 1, 2), weight=1)
        self.columnconfigure((0, 1, 2), weight=1)
        self.setup()

    def setup(self):
        """Shape the menu component"""
        # Components
        self.lbl_path_to_invs = tk.Label(
            self, text="Nie wybrano folderu", fg="grey")
        self.btn_choose_dir = tk.Button(
            self, text="Wybierz folder", command=self.get_path)
        self.lbl_dates = tk.Label(
            self, text="Przedział kursu EUR"
        )
        self.eur_date_start = DateEntry(self, date_pattern="yyyy-mm-dd")
        self.eur_date_end = DateEntry(self, date_pattern="yyyy-mm-dd")
        self.btn_gen_inv = tk.Button(
            self, text="Generuj pliki", state="disabled",
            command=self.generate_xml_and_clients
        )
        # Placement
        self.lbl_path_to_invs.grid(
            row=0, column=0, sticky="w", padx=10, pady=10)
        self.btn_choose_dir.grid(
            row=0, column=2, sticky="ew", padx=10, pady=10)
        self.lbl_dates.grid(
            row=1, column=0, sticky="w", padx=10, pady=10
        )
        self.eur_date_start.grid(
            row=1, column=1, sticky="ew", padx=10, pady=10
        )
        self.eur_date_end.grid(
            row=1, column=2, sticky="ew", padx=10, pady=10
        )
        self.btn_gen_inv.grid(
            row=2, column=2, sticky="ew", padx=10, pady=10
        )

    def get_path(self):
        """Open a window searching for workspace directory and update
        the label."""
        filepath = askdirectory()

        if not filepath:
            return

        self.lbl_path_to_invs["text"] = f"{filepath}"
        self.btn_gen_inv["state"] = "active"

    def generate_xml_and_clients(self):
        """Create xml file with invoices data extracted from submitted Excel
        files and xls file with client data."""
        # idea: add config with paths to companies for multi company export
        path = pathlib.Path(self.lbl_path_to_invs["text"])

        # load rates and configs
        exchange = cfg.get_eur_exchange_rate_nbp(
            pd.to_datetime(self.eur_date_start.get()),
            pd.to_datetime(self.eur_date_end.get())
        )
        companies = cfg.load_companies()
        holidays = cfg.load_holidays()

        code = ""
        # get the company code
        for i in range(len(companies)):
            if companies[i]["name"].casefold() in str(path.resolve()):
                code = companies[i]["id"]
                break

        exporter = Invoices(
            company_code=code,
            data_path=str(path.resolve()),
            exchange_rates=exchange,
            holidays=holidays
        )
        exporter.verify_data()
        exporter.gen_xml_layout()

        # print(exchange)
        # print(holidays)
        # print(companies)
        # print(path)

        
