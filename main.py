from request import Request
from form_output import Form_Output
from read_input import Read_input



new_request = Request()
new_form = Form_Output()


#new_form.sheet1['A4'].value = "test"
#print(new_form.sheet1['A28'].value)

#new_form.wb.save("excel_files/new.xlsx")

print(list(new_request.material_dict.keys())[4])