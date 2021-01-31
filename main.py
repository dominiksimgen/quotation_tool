from request import Request
from form_output import Form_Output

test_material_list = [123,456,789]

new_request = Request(test_material_list)
new_form = Form_Output()

print(new_request.destination)
#new_form.sheet1['A4'].value = "test"
#print(new_form.sheet1['A28'].value)

new_form.wb.save("excel_files/new.xlsx")