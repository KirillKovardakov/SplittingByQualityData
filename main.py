from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox as mb
from tkinter import Toplevel
import algorithm as alg


def func_help():
    Label(Toplevel(root), justify=LEFT, font=('Arial', 14), text='''
    Для начала работы с программой необходимо указать дату по которой будет разбиение на отдельные листы. Лучше 
    всего заранее посмотреть в самом файле как дата записана (С учётом знаков табуляции. Также можно не указывать
    дату и тогда данные просто разобьются по статусам качества).
    Затем добавить файл имеющий расширение ".csv", ".xlsx" или ".txt"(с разделителем = ';'), нажав кнопку "Открыть файл".
    Если файл описан корректно, то появится уведомление об успешной обработке. Файл сохраняется в том же месте,
    откуда брали исходный и пометкой в наименовании "_new".
    
    !!!Последующие обработки одного и тоже файла будут перезаписывать предыдущий обработанный файл!!!
    ''').grid()


def open_file():
    filepath = filedialog.askopenfilename()  # Запрос на открытие файла
    if filepath == "":  # Если путь пустой, то вывод ошибки
        mb.showerror(
            "Ошибка",
            "Необходимо выбрать файл")
    # if checkSetName.get() == 1:
    #     filepath = alg.set_name(filepath, entryName.get())
    if filepath[-3::] == "csv":  # Проверка файла на расширение CSV
        alg.filepathxl = filepath[:-4:] + '_new.xlsx'
        alg.open_filexl(alg.csv_to_excel(filepath), entryData.get())
        mb.showinfo(
            "Info",
            "Completed successfully!")
    elif filepath[-4::] == "xlsx":  # Проверка файла на расширение EXEL
        alg.filepathxl = filepath[:-5:] + '_new.xlsx'
        alg.open_filexl(filepath, entryData.get())
        mb.showinfo(
            "Info",
            "Completed successfully!")
    elif filepath[-3::] == "txt":  # Проверка файла на расширение Text
        alg.filepathxl = filepath[:-4:] + '_new.xlsx'
        alg.open_filexl(alg.txt_to_exel(filepath), entryData.get())
        mb.showinfo(
            "Info",
            "Completed successfully!")
    else:
        mb.showerror(
            "Формат файла не поддерживается",
            "Формат файла должен быть .xlsx, .csv или .txt!")


root = Tk()
root.title("SplittingByQuality")
root.geometry('500x300')

frm = ttk.Frame(root, padding=10)
frm.grid()

entryData = Entry(root)
entryData.grid(row=0, column=1)
Button(root, text="Quit", command=root.destroy).grid(column=1, row=2)
Label(root, text='Введите дату:').grid(column=0, row=0)

open_button = Button(root, text="Открыть файл", command=open_file)
open_button.grid(column=1, row=1, padx=10)

help_button = Button(root, text='?', command=func_help)
help_button.grid(column=2, row=0, padx=50, sticky='NE')

# checkSetName = IntVar()
# check_button = Checkbutton(root, text="Указать другое имя файла", variable=checkSetName, onvalue=1, offvalue=0)
# check_button.grid(row=1, column=0, sticky=NE, pady=(30, 0), padx=30)
#
# entryName = Entry(root)
# entryName.grid(row=2, column=0)
root.mainloop()
