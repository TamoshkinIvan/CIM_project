import pandas as pd
from image_insert import image_insert
import graficProc as gp
import tkinter
from tkinter import filedialog
import os


if __name__ == '__main__':
    root = tkinter.Tk()
    root.withdraw()
    currdir = os.getcwd()
    tempdir = filedialog.askopenfilename(parent=root, initialdir=currdir, title='')
    if len(tempdir) > 0:
        print(f"Текущая папка {tempdir}")

    list_equipment = gp.make_sechen_list(gp.make_max_date(tempdir))
    n = 1
    for sech in list_equipment:
        print(sech)
        itog_graf = gp.make_df_for_plotting(sech, tempdir).drop_duplicates()
        gp.graf_plotty(itog_graf)
        with pd.ExcelWriter('example.xlsx',  mode="a", engine="openpyxl") as writer:
            itog_graf.to_excel(writer, f"Sheet {n}")

        image_insert('example.xlsx', 'graf.jpg', f"Sheet {n}")
        n += 1

