import tkinter as tk

from gui.main_window import MainWindow


class SuperImporterApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.main_window = MainWindow()


if __name__ == '__main__':
    app = SuperImporterApp()
    app.mainloop()
