import os
import threading
from tkinter import messagebox

from openpyxl import load_workbook
from datetime import date

from tkinter import *

root = Tk()
root.geometry('250x100')
root.iconbitmap(r'MC.ico')
root.resizable(False, False)

__salary = StringVar()

wb = load_workbook('MoneyControl.xlsx')
wb.iso_dates = True
ws = wb.active

MAX_ROW = ws.max_row

FORMULAS = [
    date.today().strftime('%d-%m-%Y'),
    f'=(A{MAX_ROW + 1}-A{MAX_ROW})',
    f'=ROUND((C{MAX_ROW + 1}/A{MAX_ROW})*100,0)',
    f'=ROUNDDOWN((E{MAX_ROW}+(C{MAX_ROW + 1}*0.2)),0)',
    f'=ROUNDDOWN((F{MAX_ROW}+(C{MAX_ROW + 1}*0.3)),0)',
    f'=ROUNDUP((G{MAX_ROW}+(C{MAX_ROW + 1}*0.5)),0)',
    'black',
    f'=ROUNDUP(G{MAX_ROW + 1}*0.6,0)',
    f'=ROUNDDOWN(G{MAX_ROW + 1}*0.4,0)',
    'black',
    f'=ROUNDUP(I{MAX_ROW + 1}*0.4,0)',
    f'=ROUNDUP(I{MAX_ROW + 1}*0.25,0)',
    f'=ROUNDDOWN(I{MAX_ROW + 1}*0.2,0)',
    f'=ROUNDDOWN(I{MAX_ROW + 1}*0.15,0)',
]


def submit():
    inputData = __salary.get()
    if not inputData.isdigit():
        messagebox.showerror("Error", "Please enter valid in-hand salary")
        entry.delete(0, 'end')
        return

    dataToBeAdded = [int(inputData), ]
    dataToBeAdded = dataToBeAdded + FORMULAS

    print(dataToBeAdded)

    ws.append(dataToBeAdded)

    wb.save('MoneyControl.xlsx')

    wb.close()
    root.destroy()

    os.startfile('MoneyControl.xlsx')


Label(root, text="Enter in-hand salary").pack(side='top')
entry = Entry(root, textvariable=__salary, justify='center', bd=3, font='18')
entry.pack(expand=True, fill='both', side='top')
Button(root, text='Submit', command=submit).pack()

root.mainloop()
