import openpyxl, os


class Form_Output:
    def __init__(self,form_template):
        self.here = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.here)
        self.wb = openpyxl.load_workbook(f"excel_files/{form_template}")
        self.sheet1 = self.wb[self.wb.sheetnames[0]]

    def fill_form(self, request):
        self.sheet1['B11'].value = request.reference_No
        self.sheet1['H11'].value = request.requester_name
        self.sheet1['H12'].value = request.requester_contacts
        self.sheet1['B50'].value = request.stackable
        self.sheet1['B52'].value = request.dangerous_goods
        self.sheet1['B56'].value = request.li_ion
        self.sheet1['H68'].value = request.ready_date
        self.sheet1['B59'].value = request.incoterm
        self.sheet1['D59'].value = request.destination_code
        self.sheet1['B78'].value = request.shipper_code
        self.sheet1['B84'].value = request.destination_code
        

    def save_output_to_excel(self, reference_No):
        self.wb.save(f"output/Ref_{reference_No}_PG_Combined Spot Quotation  Booking Form.xlsx") 