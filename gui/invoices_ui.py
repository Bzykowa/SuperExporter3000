import tkinter as tk
from tkinter.filedialog import askopenfilename


class InvoicesUI(tk.Frame):
    """The part of the application responsible for setting up invoices
    import."""

    def __init__(self, master=None, controller=None):
        super().__init__(master)
        self.controller = controller
        self.rowconfigure((0), weight=1)
        self.columnconfigure((0, 1), weight=1)
        self.setup()

    def setup(self):
        """Shape the menu component"""
        # Components
        self.lbl_path_to_invs = tk.Label(
            self, text="Nie wybrano pliku", fg="grey")
        self.btn_choose_file = tk.Button(
            self, text="Wybierz plik", command=self.get_path)
        # Placement
        self.lbl_path_to_invs.grid(
            row=0, column=0, sticky="w", padx=10, pady=10)
        self.btn_choose_file.grid(
            row=0, column=1, sticky="ew", padx=10, pady=10)

    def get_path(self):
        """Open a window searching for an Excel file and update
        the label."""
        filepath = askopenfilename(
            filetypes=[("Excel Spreadsheet", "*.xls"),
                       ("Excel Spreadsheet", "*.xlsx*")]
        )

        if not filepath:
            return

        self.lbl_path_to_invs["text"] = f"{filepath}"
