#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function
import io
import win32com.client
from tkinter import *


window = Tk()
window.title("САГА - ПУТЕВОЧНИКИ")
window.minsize(width = 400, height=300)

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u'D:\\Музыка\\Лагеря\\Солнечный\\путевки\\test.xls')
sheet = wb.ActiveSheet


def start():
	d1 = dict()
	i = 1
	txt_d1 = open('kids_list.txt', 'w')
	while sheet.Cells(i,1).value != None : #Пока есть значения в 1 столбике делай...
		d1[i]=sheet.Cells(i,2).value
		i += 1
		txt_d1.write(str(i) + "\t"+str(sheet.Cells(i,2).value) + "\n")
	txt_d1.close()

def find(): #Поиск фамилии
	li_find = []
	if (fi_name.get() != "") and (fi_name.get() != " ") and (fi_name.get() != (len(fi_name.get()) * " ")): #Проверка на пустую строку
		word = format(fi_name.get())
		i = 0
		with io.open(r'kids_list.txt') as file:
		    for line in file:
		        if word.upper() in line.upper():
		            print(line, end='')
		            x = line
		            li_find.append(line)
		            i += 1
		list_find(li_find)
		if i == 0:
			print ("К сожалению ничего не найдено")
		name.configure(text=x)

def list_find(lf): 
	#scrollbar = Scrollbar(window)
	#scrollbar.pack()
	li_find_box = lf
	lis_listbox = Listbox(width=40)
	for name in li_find_box:
	    lis_listbox.insert(END, name)
	lis_listbox.grid(column=2, row=1)


def find_enter(event): #Событие для ентера
    find()

def change(): #Изминение записи
	pass

def save(): #Сохранение
	pass

def sprint(): #Печать
	pass

#Begin progamm 
start()
 
fi_name = Entry(window,width=30)
fi_name.grid(column=1, row=0)
fi_name.focus_set()
window.bind("<Return>", find_enter)
btn = Button(window, text="Найти!", command=find)
btn.grid(column=2, row=0)

print(fi_name.get() )

#val = sheet.Cells(2,2).value #получаем значение первой ячейки
#vals = [r[0].value for r in sheet.Range("B1:B2")] #получаем значения цепочки A1:A2

#print(sheet.Cells(1,2).value)

p_name = Label(window, text="Имя:", font=("Arial", 10))
p_name.grid(column=1, row=2)
name = Label(window, font=("Arial Bold", 15))
name.grid(column=2, row=2)

wb.Save() #сохраняем рабочую книгу
wb.Close() #закрываем ее
Excel.Quit() #закрываем COM объект

window.mainloop()
