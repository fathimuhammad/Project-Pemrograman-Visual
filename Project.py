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
 
