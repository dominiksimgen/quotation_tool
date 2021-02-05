import json, os

class Location_Masterdata:
    def __init__(self):
        self.here = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.here)
        self.f = open("database/test_locations.json", "r")
        self.data = json.load(self.f)
        self.f.close()

    def get_index_from_plant_code(self, plant_code):
        for i in range (0, len(self.data['locations'])):
            if self.data['locations'][i]['Plant Code'] == plant_code:
                return i
    
    def get_location_object(self, plant_code):
        index = self.get_index_from_plant_code(plant_code)
        if index is not None:
            return self.data['locations'][index]
        else:
            raise Exception('\n\nCheck input. Plantcode(s) not found.\nIf input correct, make sure that plant is maintained in "database/test_locations.json"\n')
        