import openpyxl
from openpyxl import *
import tkinter as tk


def addtoexcel(namabuku, penulis, penerbit, tahunterbit):
    wb = load_workbook(filename='D:\Kuliah\Semester 3\Pemrograman Visual\Project\database.xlsx')
    wb.active
    sheet = wb['InputBuku']
    wb.save(filename='D:\Kuliah\Semester 3\Pemrograman Visual\Project\database.xlsx')
    x = 2
    while True:
        if sheet['A'+ str(x)].value == None:
            sheet['A'+ str(x)].value = namabuku
            sheet['B'+ str(x)].value = penulis
            sheet['C'+ str(x)].value = penerbit
            
            sheet['D'+ str(x)].value = tahunterbit
            wb.save(filename='D:\Kuliah\Semester 3\Pemrograman Visual\Project\database.xlsx')
            break
        x+=1

def outofexcel(namabuku, penulis, penerbit, tahunterbit):
    wb2 = load_workbook(filename='D:\Kuliah\Semester 3\Pemrograman Visual\Project\database.xlsx')
    wb2.active
    sheet = wb2['OutputBuku']
    wb2.save(filename='D:\Kuliah\Semester 3\Pemrograman Visual\Project\database.xlsx')
    y = 2
    while True:
        if sheet['A'+ str(y)].value == None:
            sheet['A'+ str(y)].value = namabuku
            sheet['B'+ str(y)].value = penulis
            sheet['C'+ str(y)].value = penerbit
            sheet['D'+ str(y)].value = tahunterbit
            wb2.save(filename='D:\Kuliah\Semester 3\Pemrograman Visual\Project\database.xlsx')
            break
        y+=1

def start():    
    root = tk.Tk()
    root.title("Inventaris Buku")
    tk.Label(root, text='Nama Buku').pack(side=tk.LEFT, padx=5, pady=5)
    nama = tk.Entry(root)
    nama.pack(side=tk.LEFT, padx=5, pady=5)
    tk.Label(root, text='Penulis').pack(side=tk.LEFT, padx=5, pady=5)
    penulis = tk.Entry(root)
    penulis.pack(side=tk.LEFT, padx=5, pady=5)
    tk.Label(root, text='Penerbit').pack(side=tk.LEFT, padx=5, pady=5)
    penerbit = tk.Entry(root)
    penerbit.pack(side=tk.LEFT, padx=5, pady=5)
    tk.Label(root, text='Tahun Terbit').pack(side=tk.LEFT, padx=5, pady=5)
    tahunterbit = tk.Entry(root)
    tahunterbit.pack(side=tk.LEFT, padx=5, pady=5)
    b1 = tk.Button(root, text='In', command=lambda:[(addtoexcel(nama.get(), penulis.get(), penerbit.get(), tahunterbit.get())), (print('Success'))])
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    b2 = tk.Button(root, text='Out', command=lambda:[(outofexcel(nama.get(), penulis.get(), penerbit.get(), tahunterbit.get())), (print('Success'))])
    b2.pack(side=tk.LEFT, padx=5, pady=5)
    b3 = tk.Button(root, text='Quit', command=root.quit)
    b3.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()

start()
