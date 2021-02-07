This python program automatically populates a quotation form for airfreight.
It takes the input from input.xlsx, retrieves logistical data from SAP APO and creates a new excel form as output.
Location specific master data (incoterm, address, plantcode, etc.) is read from database/locations.json, which needs to be updated in case locations are added or changed. 