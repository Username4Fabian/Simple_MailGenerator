import pandas as pd
import requests
import json
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def verify_email():
    df = pd.read_excel('emails.xlsx', header=None)
    
    # Initialize a new Workbook
    wb = Workbook()
    ws = wb.active

    for email in df[0]:
        url = f"http://apilayer.net/api/check?access_key=a36f245f96a6&email={email}"
        response = requests.get(url)
        data = json.loads(response.text)

        # If the smtp_check is true, write the email to the Excel file
        if data.get('smtp_check', False):
            ws.append([email])
            wb.save('valid_emails.xlsx')  # Save the changes after each valid email

verify_email()
