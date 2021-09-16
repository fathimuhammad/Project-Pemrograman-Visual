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
