import pandas as pd
import openpyxl
from datetime import date
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate
import os


# os.startfile("F:/7th Sem/RTO Project/RTO-Final/OUTPUT/23BH0122A.docx", 'print') #Print file
# d = GetPrinter(yourPrinter, 2)
# print(d.keys())
class ngs:

    def finalPrint():
        base_dir = Path(__file__).parent
        word = base_dir / "Book1.docx"
        output = base_dir / "OUTPUT"
        excel = base_dir / "print_details.xlsx"
        output.mkdir(exist_ok=True)

        fd = pd.read_excel(excel, 'Sheet1')
        print(fd.head())

        for record in fd.to_dict("records"):
            doc = DocxTemplate(word)
            doc.render(record)
            output_path = output / f"{record['Registration']}.docx"
            doc.save(output_path)
            os.startfile(output / f"{record['Registration']}.docx", 'print')
        print("Printed all the docs Successfully")
