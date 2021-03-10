#version: 1.0.2

from request import Request
from excel_interface import Form_Output
import sys

def Main():
    if not sys.warnoptions: # hides warning messages from end user
        import warnings
        warnings.simplefilter("ignore")



    print("\ncreating quotation form...\n")
    form_template = "PG_Combined Spot Quotation  Booking Form_v 1_9_unprotected.xlsx"
    location_database = "locations.json"
    new_request = Request(location_database)
    new_form = Form_Output(form_template)
    new_form.fill_form(new_request)
    new_form.save_output_to_excel(new_request.reference_ID)
    print(f"\nNew quotation form saved with reference:\n{new_request.reference_ID}\n")



if __name__ == "__main__":
    Main()
    input("Press 'Enter' key to exit.")