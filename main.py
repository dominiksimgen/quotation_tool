from request import Request
from form_output import Form_Output
from read_input import Read_input



form_template = "PG_Combined Spot Quotation  Booking Form_v 1_9_uprotected.xlsx"
new_request = Request()
new_form = Form_Output(form_template)
new_form.save_output_to_excel(new_request.reference_No)

for i in range(0, len(new_request.material_object_array)):
    print(vars(new_request.material_object_array[i]))