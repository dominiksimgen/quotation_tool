
import material, functions, datetime

class Request:
    def __init__(self, material_list):
        self.reference_ID = "TestNo_0101_1999"
        self.material_array = []
        self.requester_name = "Planner Name"
        self.requester_contacts = "0049 1111 999 / planner@company.com"
        for i in material_list:
            self.material_array.append(material.Material(i))
        self.stackable = "No"
        self.dangerous_goods = "Yes"
        self.li_ion = "Yes"
        self.ready_date = functions.date_by_adding_business_days(datetime.date.today(),5)
        self.destination = "A999"





