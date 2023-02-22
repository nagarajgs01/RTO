import pandas as pd
import openpyxl
from datetime import date
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate

base_dir = Path(__file__).parent
word = "F:/7th Sem/RTO Project/Book1.docx"
excel = "F:/7th Sem/RTO Project/Book.xlsx"  # file
output = base_dir / "OUTPUT"

output.mkdir(exist_ok=True)

fd = pd.read_excel(excel, 'Sheet1')
date = date.today()
d = date.strftime("%d-%m-%y")
fd.insert(4, column="date", value=d)
fd.insert(0, column="Sno", value=range(1, len(fd) + 1))
print(fd.head())
print(fd.to_dict("records"))

for record in fd.to_dict("records"):
    if 2 <= record['Sno'] <= 4:
        doc = DocxTemplate(word)
        doc.render(record)
        output_path = output / f"{record['Name']}-doc.docx"
        doc.save(output_path)

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

