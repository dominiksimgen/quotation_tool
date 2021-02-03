import json, os

here = os.path.dirname(os.path.abspath(__file__))
os.chdir(here)
f = open("database/locations.json", "r")
data = json.load(f)
f.close()



#print(data['locations'][0]['Plant_Code'])
#print(data['locations'][1])

def get_object_from_plant_code(my_JSON, plant_code):
    for i in range (0, len(my_JSON['locations'])):
        if my_JSON['locations'][i]['Plant Code'] == plant_code:
            return i

p_codes = ["1707","7035","B777","3662","0371","2518","9774","B260","0114","8067","C595","A673","8513","5740","9716","4773","6213","C272"]
for i in range(0, len(p_codes)):
    data['locations'][i+1]['Plant Code'] = p_codes[i]

print(get_object_from_plant_code(data, '4552'))
print(len(data['locations']))

f = open("database/locations.json", "w")
json.dump(data, f)
f.close()




