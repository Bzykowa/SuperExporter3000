import tkinter as tk
import pandas as pd
import config.utils as cfg
import pathlib
import pyexcel
from tkinter.filedialog import askdirectory
from tkinter.ttk import Combobox
from tkcalendar import DateEntry

from xml_parsing.invoices import Invoices


class InvoicesUI(tk.Frame):
    """The part of the application responsible for setting up invoices
    import."""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.controller = controller
        self.rowconfigure((0, 1, 2, 3, 4), weight=1)
        self.columnconfigure((0, 1, 2), weight=1)
        self.setup()

    def setup(self):
        """Shape the menu component"""
        # month choice setup
        months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
        self.month = tk.IntVar(self)
        self.month.set(1)
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
        self.btn_check_inv = tk.Button(
            self, text="Sprawdź pliki", state="disabled",
            command=self.check_invoices
        )
        self.month_choice = Combobox(self, values=months)
        self.lbl_month = tk.Label(self, text="Wybierz miesiąc")
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
        self.lbl_month.grid(
            row=2, column=0, sticky="w", padx=10, pady=10
        )
        self.month_choice.grid(
            row=2, column=2, sticky="ew", padx=10, pady=10
        )
        self.btn_check_inv.grid(
            row=3, column=2, sticky="ew", padx=10, pady=10
        )
        self.btn_gen_inv.grid(
            row=4, column=2, sticky="ew", padx=10, pady=10
        )

    def get_path(self):
        """Open a window searching for workspace directory and update
        the label."""
        filepath = askdirectory()

        if not filepath:
            return

        self.lbl_path_to_invs["text"] = f"{filepath}"
        self.btn_check_inv["state"] = "active"
        self.btn_gen_inv["state"] = "disabled"

    def check_invoices(self):
        """
        Load invoices and check for known errors.
        If no errors, enable file generation.

        """
        # idea: add config with paths to companies for multi company export
        path = pathlib.Path(self.lbl_path_to_invs["text"])

        # load rates and configs
        exchange = cfg.get_eur_exchange_rate_nbp(
            pd.to_datetime(self.eur_date_start.get()),
            pd.to_datetime(self.eur_date_end.get())
        )
        companies = cfg.load_companies()
        holidays = cfg.load_holidays()

        self.code = ""
        # get the company code
        for i in range(len(companies)):
            if companies[i]["name"].casefold() in str(
                    path.resolve()).casefold():
                self.code = companies[i]["id"]
                break

        self.exporter = Invoices(
            company_code=self.code,
            data_path=str(path.resolve()),
            exchange_rates=exchange,
            holidays=holidays,
            month=self.month.get()
        )
        errors = self.exporter.verify_data()

        if not errors:
            self.btn_gen_inv["state"] = "active"
        else:
            print(errors)

    def generate_xml_and_clients(self):
        """Create xml file with invoices data extracted from submitted Excel
        files and xls file with client data."""

        path = pathlib.Path(self.lbl_path_to_invs["text"])

        self.exporter.gen_xml_layout()
        self.exporter.split_xml(max_records=500)
        output = self.exporter.formatted_print()
        if isinstance(output, list):
            for idx in range(len(output)):
                output_path = str(
                    path.joinpath(
                        "export_{}{}.xml".format(self.code, idx)
                    ).resolve()
                )
                with open(output_path, "wb") as out:
                    out.write(output[idx].encode('utf-8'))
        elif isinstance(output, str):
            output_path = str(
                path.joinpath(
                    "export_{}.xml".format(self.code)
                ).resolve()
            )
            with open(output_path, "wb") as out:
                out.write(output.encode('utf-8'))

        # bullshit to export df to .xls
        clients = self.exporter.get_clients_data()
        csv = str(path.joinpath("kontrahenci.csv").resolve())
        xls = str(path.joinpath("Kontrahenci.xls").resolve())
        clients.to_csv(csv, columns=[
            "Kod", "Nazwa", "Nazwa2", "Nazwa3", "Telefon", "Telefon2",
            "TelefonSms", "Fax", "Ulica", "NrDomu", "NrLokalu",
            "KodPocztowy", "Poczta", "Miasto", "Kraj", "Wojewodztwo",
            "Powiat", "Gmina", "URL", "Grupa", "OsobaFizyczna", "NIP",
            "NIPKraj", "Zezwolenie", "Regon", "Pesel", "Email",
            "BankRachunekNr", "BankNazwa", "Osoba", "Opis", "Rodzaj",
            "PlatnikVAT", "PodatnikVatCzynny", "Eksport", "LimitKredytu",
            "Termin", "FormaPlatnosci", "Ceny", "CenyNazwa", "Upust",
            "NieNaliczajOdsetek", "MetodaKasowa", "WindykacjaEMail",
            "WindykacjaTelefonSms", "AlgorytmNettoBrutto", "Waluta"
        ], index=False)
        pyexcel.save_as(file_name=csv, dest_file_name=xls)
