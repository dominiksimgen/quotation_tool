from request import Request
from form_output import Form_Output




form_template = "PG_Combined Spot Quotation  Booking Form_v 1_9_unprotected.xlsx"
new_request = Request()
new_form = Form_Output(form_template)
new_form.fill_form(new_request)
new_form.save_output_to_excel(new_request.reference_No)