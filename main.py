from request import Request
from form_output import Form_Output
import sys

if not sys.warnoptions:
    import warnings
    warnings.simplefilter("ignore")




form_template = "PG_Combined Spot Quotation  Booking Form_v 1_9_unprotected.xlsx"
location_database = "test_locations.json"
new_request = Request(location_database)
new_form = Form_Output(form_template)
new_form.fill_form(new_request)
new_form.save_output_to_excel(new_request.reference_ID)
print(f"\nNew quotation form saved with reference:\n{new_request.reference_ID}\n")