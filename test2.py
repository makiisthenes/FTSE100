import datetime
import shutil
import openpyxl
from openpyxl import load_workbook
now = datetime.datetime.now()
yearmonth = now.strftime("%b_%Y")
print(yearmonth)
path = r"./test/report_{}.xlsx".format(yearmonth)
shutil.copy(r"./report.xlsx",path)
print(path)
report = r"./test/ftse100.xlsx"
wb = load_workbook(report)
wb.create_sheet("sheet1")
wb.save(report)
