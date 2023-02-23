import pandas as pd
import openpyxl
from datetime import date
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate


def comma(a):
    str = []
    for i in a:
        if i.__contains__("-"):
            hifen(i)
        else:
            for record in fd.to_dict("records"):
                c = int(i)
                if record['Sno'] == c:
                    str.append(c)
                    doc = DocxTemplate(word)
                    doc.render(record)
                    output_path = output / f"{record['Name']}-doc.docx"
                    doc.save(output_path)
    if len(str) > 0:
        print(f"Successfully printed data with sno : {str}")


def hifen(b):
    lst = b.split('-')
    print(lst)
    a = int(lst[0])
    b = int(lst[1])
    for record in fd.to_dict("records"):
        if a <= record['Sno'] <= b:
            doc = DocxTemplate(word)
            doc.render(record)
            output_path = output / f"{record['Name']}-doc.docx"
            doc.save(output_path)
    print(f"Successfully printed from {a} to {b}")


base_dir = Path(__file__).parent
word = "F:/7th Sem/RTO Project/Book1.docx"
excel = "F:/7th Sem/RTO Project/Book.xlsx"  # file
output = base_dir / "OUTPUT"

output.mkdir(exist_ok=True)

# fd = pd.read_excel(excel, 'Sheet1')
# date = date.today()
# d = date.strftime("%d-%m-%y")
# fd.insert(4, column="date", value=d)
# fd.insert(0, column="Sno", value=range(1, len(fd) + 1))
# print(fd.head())
# print(fd.to_dict("records"))

fd = pd.read_excel(excel, 'Sheet1')
date = date.today()
d = date.strftime("%d-%m-%y")
fd.insert(4, column="date", value=d)
fd.insert(0, column="Sno", value=range(1, len(fd) + 1))
print(fd.head())
print(fd.to_dict("records"))

print("1 ) If want to print a range \n2 ) Print All")
ch = input("Enter your choice")
print(ch)

if ch == '1':
    range = input("Enter the range")
    if range.__contains__(","):
        ele = range.split(",")
        comma(ele)
    elif range.__contains__("-"):
        hifen(range)
else:
    for record in fd.to_dict("records"):
        doc = DocxTemplate(word)
        doc.render(record)
        output_path = output / f"{record['Name']}-doc.docx"
        doc.save(output_path)
    print("Printed all the docs Successfully")

# if (ch == '1'):
#     sno = int(input("Enter the sno of the member"))
#     for record in fd.to_dict("records"):
#         if record['Sno'] == sno:
#             doc = DocxTemplate(word)
#             doc.render(record)
#             output_path = output / f"{record['Name']}-doc.docx"
#             doc.save(output_path)
#     print("Printed document with sno")
#
# elif (ch == '2'):
#     range = input("Enter the Range seperated by - :ex 1-4 ")
#     lst = range.split('-')
#     print(lst)
#     a = int(lst[0])
#     b = int(lst[1])
#     for record in fd.to_dict("records"):
#         if a <= record['Sno'] <= b:
#             doc = DocxTemplate(word)
#             doc.render(record)
#             output_path = output / f"{record['Name']}-doc.docx"
#             doc.save(output_path)
#     print("Printed docs Succesfully")
# else:
#     for record in fd.to_dict("records"):
#         doc = DocxTemplate(word)
#         doc.render(record)
#         output_path = output / f"{record['Name']}-doc.docx"
#         doc.save(output_path)
#     print("Printed all the docs Successfully")

# if record['Done'].equals("True"):


# names = fd["names"].values.tolist()
# div = fd['div'].values.tolist()
# count = 0
#
# while count < len(names):
#     print(f"Name : {names[count]} \n Div : {div[count]}")
#     count += 1
#
# pd.
