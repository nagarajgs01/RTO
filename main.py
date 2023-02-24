import pandas as pd
import openpyxl
from datetime import date
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate
from tkinter import messagebox


class ngs:
    global output
    global word
    global excel
    global fd

    def getBasePath():
        base_dir = Path(__file__).parent
        return base_dir

    def pathAloc(base_dir):
        ngs.word = base_dir / "Book1.docx"
        ngs.output = base_dir / "OUTPUT"

    def xlPath(xlpt):
        ngs.excel = xlpt  # file

    def createDir(output):
        output.mkdir(exist_ok=True)

    def comma(a):
        str = []
        for i in a:
            if i.__contains__("-"):
                ngs.hifen(i)
            else:
                for record in ngs.fd.to_dict("records"):
                    c = int(i)-1
                    if record['Sno'] == c:
                        str.append(c)
                        doc = DocxTemplate(ngs.word)
                        doc.render(record)
                        output_path = ngs.output / f"{record['Name']}-doc.docx"
                        doc.save(output_path)
        if len(str) > 0:
            print(f"Successfully printed data with sno : {str}")

    def hifen(c):
        lst = c.split('-')
        print(lst)
        a = int(lst[0])-1
        b = int(lst[1])-1
        for record in ngs.fd.to_dict("records"):
            if a <= record['Sno'] <= b:
                doc = DocxTemplate(ngs.word)
                doc.render(record)
                output_path = ngs.output / f"{record['Name']}-doc.docx"
                doc.save(output_path)
        print(f"Successfully printed from {a} to {b}")
        # messagebox.showinfo("Success", f"Successfully printed from {a} to {b}")

    # fd = pd.read_excel(excel, 'Sheet1')
    # date = date.today()
    # d = date.strftime("%d-%m-%y")
    # fd.insert(4, column="date", value=d)
    # fd.insert(0, column="Sno", value=range(1, len(fd) + 1))
    # print(fd.head())
    # print(fd.to_dict("records"))

    def readXL(xlPath):
        ngs.fd = pd.read_excel(xlPath, 'Sheet1')
        # date = date.today()
        # d = date.strftime("%d-%m-%y")
        # fd.insert(4, column="date", value=d)
        fd=ngs.fd
        fd.insert(0, column="Sno", value=range(1, len(fd) + 1))
        print(fd.head())
        print(fd.to_dict("records"))
        # return fd

    '''
    print("1 ) If want to print a range \n2 ) Print All")
    ch = input("Enter your choice")
    print(ch)
    
    '''

    def rngSelected(range):
        # range = input("Enter the range : ")
        if range.__contains__(","):
            ele = range.split(",")
            ngs.comma(ele)
        elif range.__contains__("-"):
            ngs.hifen(range)
        else:
            for record in ngs.fd.to_dict("records"):
                rng = int(range)
                if rng <= len(ngs.fd):
                    if record['Sno'] == rng-1:
                        doc = DocxTemplate(ngs.word)
                        doc.render(record)
                        output_path = ngs.output / f"{record['Name']}-doc.docx"
                        doc.save(output_path)
                    print(f"Printed document with sno : {rng}")
                    # messagebox.showinfo("Success",f"Printed document with sno : {rng}")
                else:
                    print(f"{rng} not present")
                    # messagebox.showerror("Error",f"{rng} not present")

    def allSelected():
        for record in ngs.fd.to_dict("records"):
            doc = DocxTemplate(ngs.word)
            doc.render(record)
            output_path = ngs.output / f"{record['Name']}-doc.docx"
            doc.save(output_path)
        print("Printed all the docs Successfully")


# pt = ngs.getBasePath()
# ngs.pathAloc(pt)
# ngs.createDir(ngs.output)
# ngs.readXL(ngs.excel)
# ch = input("Enter Choice :")
# if ch == "1":
#     ngs.rngSelected()
# elif ch == "2":
#     ngs.allSelected()
# else:
#     print("Invalid Choice")

