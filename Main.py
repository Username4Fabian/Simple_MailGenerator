import pandas as pd
import openpyxl
import os

# Email presets
email_formats = ["{f}.{l}@example.com", "{f}{l}@example.com", "{f}_{l}@example.com"]

if(os.path.exists("emails.xlsx") == False):
    filename = 'emails.xlsx'
    workbook = openpyxl.Workbook()
    workbook.save(filename)

# Function to generate email addresses
def generate_emails(first_name, last_name):
    emails = []
    for email_format in email_formats:
        email = email_format.format(f=first_name, l=last_name)
        emails.append(email)
    return emails

# Function to get names from user
def get_names():
    names = []
    line_counter = 1  # Initialize line counter

    while True:
        print("Enter a name (or 'done' to finish):")
        full_name = input("Full name: ").strip()
        if full_name.lower() == 'done':
            break
        first_name, last_name = full_name.split(maxsplit=1)
        names.append((first_name, last_name))
        
        # Load the workbook
        workbook = openpyxl.load_workbook('emails.xlsx')

        # Select the first sheet (you can modify this if you have multiple sheets)
        sheet = workbook.active

        # Find the last line counted in column A
        while sheet.cell(row=line_counter, column=1).value is not None:
            line_counter += 1

        # Generate emails for the current name
        all_emails = generate_emails(first_name, last_name)

        # Write the emails to the next line
        for email in all_emails:
            sheet.cell(row=line_counter, column=1).value = email

            # Increment the line counter
            line_counter += 1

        # Save the workbook
        workbook.save('emails.xlsx')

    return names

# Get names from user
names = get_names()
