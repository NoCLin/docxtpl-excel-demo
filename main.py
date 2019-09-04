# coding:utf-8

import pandas as pd
from docxtpl import DocxTemplate

FROM_EXCEL = 'data.xls'
SHEET_NAME = 'Sheet1'
TEMPLATE_DOCX = "template.docx"
OUTPUT = "generated.docx"

df = pd.read_excel(FROM_EXCEL, sheet_name=SHEET_NAME)

rows = []
for row_arr in df.iterrows():
    row = {col: row_arr[1][col] for col in df}
    rows.append(row)
print(rows)
doc = DocxTemplate(TEMPLATE_DOCX)
context = {"rows": rows, }
doc.render(context)
doc.save(OUTPUT)
