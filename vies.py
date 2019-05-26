import datetime as dt
import requests
import pandas as pd
from bs4 import BeautifulSoup


input_file = pd.read_excel("vat.xlsx")
# dictionary for storing results
result_data = {"VAT_no": [], "valid": [], "time_stamp": []}

# iterate through all vat numbers in input_file
for vat_no in input_file["DIČ"]:

    # split the whole VAT no to state code and number
    state_code = str(vat_no).strip()[:2]
    number = str(vat_no).strip()[2:]

    # url address to be scraped
    url = "http://ec.europa.eu/taxation_customs/vies/vatResponse.html"
    payload = {"memberStateCode": state_code, "number": number}
    headers = {"Accept-Language": "cs"}

    try:
        page = requests.post(url, data=payload, headers=headers, timeout=5)

        # check if the page was loaded correctly
        if page.status_code == requests.codes.ok:
            bs = BeautifulSoup(page.text)

        # area of the webpage with required data
        validation_table = bs.find("table", id="vatResponseFormTable")
        result = validation_table.find("span").text
        time = validation_table.find(string="Datum přijetí žádosti").find_next().text

        # save the data do result_data dictionary
        result_data["VAT_no"].append(vat_no)
        result_data["valid"].append(result)
        result_data["time_stamp"].append(time)
        print(vat_no)
    except:
        result_data["VAT_no"].append(vat_no)
        result_data["valid"].append("timeout")
        result_data["time_stamp"].append(dt.datetime.now())
        print(vat_no)

# creage pandas DataFrame from result_data dictionary    
result_table = pd.DataFrame.from_dict(result_data)
# save the result table to excel
result_table.to_excel("vies_validation.xlsx", index=False)

