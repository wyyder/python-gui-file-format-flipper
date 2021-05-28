import csv
import openpyxl as xl

import tkinter
from tkinter import Label
from tkinter.messagebox import showinfo
import windnd


def dragged_file(files):
    for item in files:
        file_path = item.decode('gbk')
        check_file(file_path=file_path)
    pass


def check_file(file_path: str):
    if file_path.endswith(".txt"):
        convert_txt_to_excel(file_path=file_path)
        showinfo("Success", "Done!")
    elif file_path.endswith(".xlsx"):
        convert_excel_to_txt(file_path=file_path)
        showinfo("Success", "Done!")
    else:
        showinfo("Error", "Unsupported file format!")
    pass


def convert_txt_to_excel(file_path: str):
    output_file = file_path.replace('.txt', '.xlsx')
    from xlsxwriter.workbook import Workbook
    wb = Workbook(output_file)
    sheet = wb.add_worksheet()
    with open(file_path) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        for r, row in enumerate(csv_reader):
            for c, col in enumerate(row):
                sheet.write(r, c, col)
    wb.close()
    pass


def convert_excel_to_txt(file_path: str):
    wb = xl.load_workbook(file_path)
    sheet = wb.worksheets[0]
    output_file = file_path.replace('.xlsx', '.txt')
    with open(output_file, "a") as out_file:
        for row in sheet.rows:
            values = []
            for cell in row:
                values.append(str(cell.value))
            out_file.write(";".join(values) + '\n')
    pass


root = tkinter.Tk()
root.title('Wyyder - File flipper TXT / XLSX')
root.geometry("400x200")
text = Label(root, text="Drag & Drop file here to Convert").place(x=100, y=80)
windnd.hook_dropfiles(root, func=dragged_file)
root.mainloop()
