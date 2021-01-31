import cmath, openpyxl, os
import pandas as pd

class Form_Output:
    def __init__(self):
        self.here = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.here)
        self.wb = openpyxl.load_workbook("excel_files/PG_Combined Spot Quotation  Booking Form_v 1_9_uprotected.xlsx")
        self.sheet1 = self.wb[self.wb.sheetnames[0]]






    