import datetime, time
def date_by_adding_business_days(from_date, add_days):
    business_days_to_add = add_days
    current_date = from_date
    while business_days_to_add > 0:
        current_date += datetime.timedelta(days=1)
        weekday = current_date.weekday()
        if weekday >= 5: # sunday = 6
            continue
        business_days_to_add -= 1
    return current_date

def create_unique_reference(country):
    current_date = datetime.date.today()
    current_second = str(int(time.time()))[-5:]
    reference_string = f"{current_date}_{current_second}_Quotation_Form_{country}"
    return reference_string




