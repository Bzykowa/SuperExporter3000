import tkinter as tk
import pathlib
from tkinter.filedialog import askopenfilename, askdirectory
from config.utils import load_companies
from xml_parsing.delegations import Delegations


class DelegationsUI(tk.Frame):
    """The part of the application responsible for setting up delegation
    import."""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.controller = controller
        self.rowconfigure((0, 1, 2, 3), weight=1)
        self.columnconfigure((0, 1), weight=1)
        self.no_file_message = "Nie wybrano pliku/folderu"
        self.setup()

    def setup(self):
        """Shape the menu component"""
        # Components
        self.lbl_path_to_dels = tk.Label(
            self, text=self.no_file_message, fg="grey")
        self.btn_choose_file = tk.Button(
            self, text="Wybierz plik", command=self.get_file)
        self.btn_choose_dir = tk.Button(
            self, text="Wybierz folder", command=self.get_dir)
        self.lbl_file_errors = tk.Label(
            self, text=self.no_file_message, fg="grey"
        )
        self.btn_generate_file = tk.Button(
            self, text="Generuj xml",  state="disabled",
            command=self.generate_xml
        )
        # Placement
        self.lbl_path_to_dels.grid(
            row=0, column=0, sticky="w", padx=10, pady=10)
        self.btn_choose_file.grid(row=0, column=1, sticky="ew", padx=10)
        self.btn_choose_dir.grid(row=1, column=1, sticky="ew", padx=10)
        self.lbl_file_errors.grid(
            row=2, column=0, sticky="w", padx=10, pady=10)
        self.btn_generate_file.grid(
            row=3, column=1, padx=10, pady=10, sticky="ew"
        )

    def get_file(self):
        """Open a window searching for an Excel file and update
        the label."""
        filepath = askopenfilename(
            filetypes=[("Excel Spreadsheet", "*.xls"),
                       ("Excel Spreadsheet", "*.xlsx*")]
        )

        if not filepath:
            return

        self.lbl_path_to_dels["text"] = f"{filepath}"
        self.lbl_file_errors["text"] = ""
        self.btn_generate_file["state"] = "active"
        self.directory_mode = False

    def get_dir(self):
        """Open a window searching for a directory with Excel files to import
        and update the label."""
        filepath = askdirectory()

        if not filepath:
            return

        self.lbl_path_to_dels["text"] = f"{filepath}"
        self.lbl_file_errors["text"] = ""
        self.btn_generate_file["state"] = "active"
        self.directory_mode = True

    def generate_xml(self):
        """Create xml file with data extracted from submitted Excel files"""

        path = pathlib.Path(self.lbl_path_to_dels["text"])

        # load company code config
        companies = load_companies()
        # search for modern excel files if in dir mode
        files = [str(p.resolve()) for p in path.glob(
            "*.xlsx")] if self.directory_mode else [str(path.resolve())]

        # match files to company codes
        exports = [
            (companies[i]["id"], p) for i in range(
                len(companies)
            )for p in files if companies[i]["name"].casefold() in p.casefold()
        ]

        for file in exports:
            exporter = Delegations(file[0], file[1])
            exporter.gen_xml_layout()
            output_path = str(path.joinpath(
                "export_{}.xml".format(file[0])).resolve())

            with open(output_path, "wb") as output:
                output.write(exporter.formatted_print().encode('utf-8'))
