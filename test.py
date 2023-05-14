import smtplib
from email_validator import validate_email

def is_valid_email(email):
    try:
        v = validate_email(email, check_deliverability=True)
        email_address = v["email"]
        # Split the email address to extract the domain
        domain = email_address.split("@")[1]
        
        # Connect to the mail server of the email address domain
        server = smtplib.SMTP(domain)
        server.ehlo()
        server.quit()
        
        print(f"The email address '{email_address}' is valid and exists.")
        return True
    except Exception as e:
        print(f"The email address '{email}' is not valid or does not exist. Error: {str(e)}")
        return False

# Usage
testEmail = "example@stackabuse.com"
is_valid_email(testEmail)
print("test")