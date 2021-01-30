import cmath, openpyxl, os
import pandas as pd

here = os.path.dirname(os.path.abspath(__file__))
os.chdir(here)

wb = openpyxl.load_workbook("excel_files/Copy of PG_Combined Spot Quotation  Booking Form_v 1_9.xlsx")
sheet1 = wb[wb.sheetnames[0]]

print(os.getcwd())
print(wb.sheetnames)


sheet1['C28'].value = "test1"
sheet1['A28'].value = 1.45
#sheet1['D28'].value = "test2"
print(sheet1['C28'].value)
print(sheet1['D28'].value)
wb.save("excel_files/new.xlsx")








