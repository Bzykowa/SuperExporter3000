import tkinter as tk

from gui.main_window import MainWindow


class SuperImporterApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("SuperImporter3000")
        self.main_window = MainWindow()
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.main_window.grid(sticky="nsew")
        self.resizable(True, True)
        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())


if __name__ == '__main__':
    app = SuperImporterApp()
    app.mainloop()
