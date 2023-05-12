import pandas as pd
import openpyxl
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QMessageBox, QGridLayout, QSizePolicy
from PyQt5.QtGui import QPalette, QColor, QFont
from PyQt5.QtCore import Qt

#To do:
# Make only first or last name valid (done)
# Expand email formats (done)
# UI (done)

# get mail formats from json file
# validate Emails (might be possible)

# Timing: 
# 11.05.2023 20:30 - 01:00
# 12.05.2023 during day (2h)
# 12.05.2023 22:00 - 

if(os.path.exists("emails.xlsx") == False):
    filename = 'emails.xlsx'
    workbook = openpyxl.Workbook()
    workbook.save(filename)

# Email presets
email_domains = ["@gmail.com", "@yahoo.com", "@hotmail.com", "@aon.at", "@gmx.at", "@outlook.com", "@live.com", "@icloud.com"]
email_formats = []

for x in email_domains:
    email_formats += ["{f}.{l}"+x, "{f}{l}"+x, "{f}_{l}"+x, "{f[0]}.{l}"+x, "{f}.{l[0]}"+x,]


# Email presets for first name only
first_name_email_formats = ["{f}"+x for x in email_domains]

# Function to generate email addresses
def generate_emails(first_name, last_name):
    emails = []
    if last_name:
        for email_format in email_formats:
            email = email_format.format(f=first_name, l=last_name)
            emails.append(email)
    else:
        for email_format in first_name_email_formats:
            email = email_format.format(f=first_name)
            emails.append(email)
    return emails

# Function to get names from user
def get_names():
    names = []
    line_counter = 1  # Initialize line counter

    full_name = name_entry.text().strip().lower()
    if full_name == '':
        QMessageBox.critical(None, "Error", "Please enter a name")
        return

    name_parts = full_name.split(" ")
    if len(name_parts) == 2:
        first_name, last_name = name_parts
    else:
        first_name = name_parts[0]
        last_name = ""

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

    name_entry.clear() # clear the input field
    QMessageBox.information(None, "Success", f"Emails for {full_name} have been generated")

# New function to open the Excel file
def open_excel():
    filename = 'emails.xlsx'
    if os.path.isfile(filename):
        if os.name == 'nt':
            os.system('start excel.exe "%s"' % filename)
        elif os.name == 'posix':
            os.system('open "%s"' % filename)
        else:
            QMessageBox.critical(None, "Error", "OS not supported")
    else:
        QMessageBox.critical(None, "Error", "File not found")


# GUI setup
app = QApplication([])
window = QWidget()

# Set the window size and background color
window.resize(800, 600)  # Set window size to 800x600
palette = QPalette()
palette.setColor(QPalette.Window, QColor(44, 47, 51))  # Discord-like dark background color
window.setPalette(palette)

outer_layout = QVBoxLayout()  # Outer layout to center the grid layout vertically
layout = QGridLayout()  # Grid layout to center widgets horizontally

header_label = QLabel("E-Mail Generator")
header_label.setFont(QFont('Arial', 20))
header_label.setStyleSheet("color: white")
header_label.setAlignment(Qt.AlignCenter)  # Align the header to the center

name_label = QLabel("Full Name:")
name_label.setFont(QFont('Arial', 16))  # Make the "Full Name:" label a bit smaller
name_label.setStyleSheet("color: white")  # Set label text color to white like Discord
name_label.setAlignment(Qt.AlignCenter)  # Align the label to the center

name_entry = QLineEdit()
name_entry.setStyleSheet("background-color: white; color: black, size:200%")  # Set entry field background to white and text to black

generate_button = QPushButton("Generate Emails")
generate_button.setStyleSheet("background-color: #7289DA; color: white")  # Discord-like button color
generate_button.clicked.connect(get_names)
name_entry.returnPressed.connect(get_names)

# New button to open the Excel file
open_button = QPushButton("Open Excel File")
open_button.setStyleSheet("background-color: #7289DA; color: white")  # Discord-like button color
open_button.clicked.connect(open_excel)

# Place the widgets in the middle of the grid
layout.addWidget(header_label, 0, 0)
layout.addWidget(name_label, 1, 0)
layout.addWidget(name_entry, 2, 0)
layout.addWidget(generate_button, 3, 0)
layout.addWidget(open_button, 4, 0)  # Add the new button to the grid layout
layout.setVerticalSpacing(20)  

# Add the grid layout to the outer layout
outer_layout.addStretch()
outer_layout.addLayout(layout)
outer_layout.addStretch()

window.setLayout(outer_layout)
window.show()

app.exec_()
