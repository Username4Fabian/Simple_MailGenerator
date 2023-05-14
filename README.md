# E-Mail Generator

This is a Python-based email generator that creates a list of potential emails based on a given name. The script works by taking a name and applying it to several email formats, which are then saved in an Excel file. The script uses PyQt5 for a graphical user interface, so you can enter names and generate emails directly in the application.

## Prerequisites

This project requires Python 3 and the following Python libraries installed:

- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [PyQt5](https://pypi.org/project/PyQt5/)

You can install these dependencies with pip:

```bash
pip install pandas openpyxl PyQt5
```

## Running the Program

To run this program, you can simply run the Python script:

```bash
python3 email_generator.py
```

This will open up a graphical user interface. Enter a full name (first name and last name) into the text field and click "Generate Emails". This will generate potential emails and save them into an Excel file named 'emails.xlsx'. 

If you only provide a first name, the program will generate emails using only the first name.

You can also click "Open Excel File" to open the 'emails.xlsx' file directly from the application.

## Email Settings

The email formats and domains used for generating emails can be customized. You can modify the 'email_settings.txt' file in the following format:

```json
{
    "email_domains": ["@gmail.com", "@aon.at", "@gmx.at", "@gmx.net", "@outlook.com", "@icloud.com"],
    "email_format_structures": ["{f}.{l}", "{f}{l}", "{f}_{l}", "{f[0]}.{l}", "{f}.{l[0]}", "{l}{f}"]
}
```

The `{f}` and `{l}` are placeholders for the first name and last name respectively. You can adjust these structures and domains as you see fit.

If the 'email_settings.txt' file is not found, the program will use the default values and create a new 'email_settings.txt' file.

## Author

This script was created by Fabian Sykes.



