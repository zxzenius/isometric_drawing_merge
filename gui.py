import os
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askdirectory
import isogen_purge
import merge_drawings


def merge():
    merge_drawings.merge(os.getcwd())


def purge():
    isogen_purge.process(os.getcwd())


class Application(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title('Isogen Tool')
        self.master.geometry('200x200')
        self.grid(sticky=W+E+N+S)
        pad_x = 20
        self.cwd = os.getcwd()
        self.info = Label(self, text=self.cwd)
        self.select_btn = ttk.Button(self, text="Select Folder", command=self.select_folder)
        self.select_btn.grid(row=1, column=0, columnspan=1, sticky=W + E, padx=pad_x, pady=10, ipady=20, ipadx=37)
        self.purge_btn = ttk.Button(self, text="Purge", command=purge)
        self.purge_btn.grid(row=2, column=0, sticky=W + E, padx=pad_x, pady=10)
        self.merge_btn = ttk.Button(self, text="Merge", command=merge)
        self.merge_btn.grid(row=3, column=0, sticky=W + E, padx=pad_x, pady=10)
        self.info.grid(row=4, sticky=W + E, padx=5)

    def select_folder(self):
        path = askdirectory(initialdir=os.getcwd(), mustexist=True)
        if path:
            print(f'Select: "{path}"')
            os.chdir(path)
            self.info.config(text=path)


if __name__ == '__main__':
    root = Tk()
    app = Application(master=root)
    app.mainloop()
