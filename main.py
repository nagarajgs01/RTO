import pandas as pd
# import openpyxl
from datetime import date
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate

base_dir = Path(__file__).parent
word = "F:/7th Sem/RTO Project/Book1.docx"
excel = "F:/7th Sem/RTO Project/Book.xlsx"
output = base_dir / "OUTPUT"

output.mkdir(exist_ok=True)

fd = pd.read_excel(excel, 'Sheet1')
date = date.today()

for record in fd.to_dict("records"):
    doc = DocxTemplate(word)
    doc.render(record)
    output_path = output / f"{record['Name']}-doc.docx"
    doc.save(output_path)
    # print(record['Done'])

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
