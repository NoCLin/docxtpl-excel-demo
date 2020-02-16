# coding:utf-8

import xlrd
from docxtpl import DocxTemplate

FROM_EXCEL = 'data.xls'
SHEET_NAME = 'Sheet1'
TEMPLATE_DOCX = "template.docx"
OUTPUT = "generated.docx"

wb = xlrd.open_workbook(FROM_EXCEL)
sheet = wb.sheet_by_name(SHEET_NAME)

rows = [
    dict(zip(sheet.row_values(0), sheet.row_values(i)))
    for i in range(1, sheet.nrows)
]

for row in rows:
    print(row)

doc = DocxTemplate(TEMPLATE_DOCX)
context = {"rows": rows, }
doc.render(context)
doc.save(OUTPUT)
