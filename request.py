
import material, functions, datetime
from read_input import Read_input

class Request:
    def __init__(self):
        self.input = Read_input()
        self.reference_ID = "TestNo_0101_1999"
        self.requester_name = self.input.requester_name
        self.requester_contacts = self.input.requester_contacts
        self.stackable = self.input.stackable
        self.dangerous_goods = self.input.dangerous_goods
        self.li_ion = self.input.li_ion
        self.ready_date = functions.date_by_adding_business_days(datetime.date.today(),5)
        self.shipper_code = self.input.shipper_code
        self.destination_code = self.input.destination_code
        self.incoterm = "FCA"
        self.material_dict = self.input.material_dict





