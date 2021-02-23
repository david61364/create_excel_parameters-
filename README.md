# create_excel_parameters-
Writing a short program in Python to create a voip parameter template.

These two programs take a set of predefined Yealink W60B parameters, takes input such as the store number, the extensions of the two W56H handsets and the store voicemail user. The .xlsx file can then be loaded to our cloud service. Yealink bases are provisioned with ZTP. Both programs output the .xlsx files each for a load balanced W60B. The customer has several stores with 4 to 6 W56H handsets at each store. Our director wanted to load balance the handsets between 2 W60Bs and be able to check the stores voicemail from any of the handsets so I had to register the main voicemail user to both units. Because this got a little complicated to implement for a couple of hundred stores, I was looking for a way to automate the parameter building process.
