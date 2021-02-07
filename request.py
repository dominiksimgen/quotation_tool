
import functions, datetime
from excel_interface import Read_excel_input
from material import Material
from location_data import Location_Masterdata

class Request:
    def __init__(self, locations_database):
        self.input = Read_excel_input()
        self.locations = Location_Masterdata(locations_database)
        self.requester_name = self.input.requester_name
        self.requester_contacts = self.input.requester_contacts
        self.stackable = self.input.stackable
        self.dangerous_goods = self.input.dangerous_goods
        self.li_ion = self.input.li_ion
        self.ready_date = functions.date_by_adding_business_days(datetime.date.today(),5)
        #shipper information:
        self.shipper_code = self.input.shipper_code
        self.shipper_company_name = self.locations.get_location_object(self.shipper_code)['Company Name']
        self.shipper_address = self.locations.get_location_object(self.shipper_code)['Address']
        #destination information:
        self.destination_code = self.input.destination_code
        self.incoterm = self.locations.get_location_object(self.destination_code)['Incoterm']
        self.airport = self.locations.get_location_object(self.destination_code)['Airport']
        self.destination_company_name = self.locations.get_location_object(self.destination_code)['Company Name']
        self.destination_address = self.locations.get_location_object(self.destination_code)['Address']
        self.reference_ID = functions.create_unique_reference(self.locations.get_location_object(self.destination_code)['Address']['Country'].replace(" ","_"))
        #list of shipment items:
        self.material_dict = self.input.material_dict
        self.material_object_array = []
        for key, value in self.material_dict.items():
            self.material_object_array.append( Material( key, value))






