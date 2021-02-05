import json, os

class location_masterdata:
    def __init__(self):
        self.here = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.here)
        self.f = open("database/locations.json", "r")
        self.data = json.load(self.f)
        self.f.close()



def get_object_from_plant_code(my_JSON, plant_code):
    for i in range (0, len(my_JSON['locations'])):
        if my_JSON['locations'][i]['Plant Code'] == plant_code:
            return i

my_JSON_file = location_masterdata()
print(get_object_from_plant_code(my_JSON_file.data, 'C595'))





