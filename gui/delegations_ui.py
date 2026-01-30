import tkinter as tk
from tkinter.filedialog import askopenfilename
from xml_parsing.delegations import Delegations


class DelegationsUI(tk.Frame):
    """The part of the application responsible for setting up delegation
    import."""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.controller = controller
        self.rowconfigure((0, 1, 2), weight=1)
        self.columnconfigure((0, 1), weight=1)
        self.no_file_message = "Nie wybrano pliku/folderu"
        self.setup()

    def setup(self):
        """Shape the menu component"""
        # Components
        self.lbl_path_to_dels = tk.Label(
            self, text=self.no_file_message, fg="grey")
        self.btn_choose_file = tk.Button(
            self, text="Wybierz plik/folder", command=self.get_path)
        self.btn_check_file = tk.Button(
            self, text="Skanuj", command=self.check_file
        )
        self.lbl_scan_dels = tk.Label(
            self, text="Sprawdź delegacje względem najpopularniejszych błędów"
        )
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
        self.lbl_scan_dels.grid(
            row=1, column=0, sticky="w", padx=10, pady=10)
        self.btn_check_file.grid(row=1, column=1, sticky="ew", padx=10)
        self.lbl_file_errors.grid(
            row=2, column=0, sticky="w", padx=10, pady=10)
        self.btn_generate_file.grid(
            row=3, column=1, padx=10, pady=10, sticky="ew"
        )

    def get_path(self):
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

    def check_file(self):
        """Scan the chosen delegations file for common errors and list
        them for user."""
        if self.lbl_path_to_dels["text"] == self.no_file_message:
            self.lbl_file_errors["text"] = self.no_file_message
            self.btn_generate_file["state"] = "disabled"
        else:
            self.btn_generate_file["state"] = "active"

    def generate_xml(self):
        """Create xml file with data extracted from submitted excel file"""
        test_id = "AcMed"
        test_output = "C:\\SuperImporter\\SuperExporter3000\\test_files\\" + \
            "export_{}.xml".format(test_id)
        exporter = Delegations(test_id, self.lbl_path_to_dels["text"])
        exporter.gen_xml_layout()
        with open(test_output, "wb") as output:
            output.write(exporter.formatted_print().encode('utf-8'))
