import openpyxl, os


class Form_Output:
    def __init__(self,form_template):
        self.here = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.here)
        self.wb = openpyxl.load_workbook(f"excel_files/{form_template}")
        self.sheet1 = self.wb[self.wb.sheetnames[0]]

    def save_output_to_excel(self, reference_No):
        self.wb.save(f"output/Ref_{reference_No}_PG_Combined Spot Quotation  Booking Form.xlsx")






    