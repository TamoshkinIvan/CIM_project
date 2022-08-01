import openpyxl
import tkinter
from tkinter import filedialog
import os


def add_image():
    root = tkinter.Tk()
    root.withdraw()
    currdir = os.getcwd()
    tempdir = filedialog.askopenfilename(parent=root, initialdir=currdir, title='Выбери папку')
    if len(tempdir) > 0:
        print(f"Выбранная папка {tempdir}")
