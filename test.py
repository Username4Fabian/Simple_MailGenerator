import smtplib
import socket

def validate_email(email_address, mail_server):
    try:
        # Create a socket with a timeout
        server = smtplib.SMTP(mail_server, timeout=10)

        # Use the ehlo or helo command to identify yourself to the SMTP server.
        server.ehlo_or_helo_if_needed()

        # Request the server for the responses to the RCPT TO command
        response = server.verify(email_address)

        # Close the connection to the server
        server.quit()

        # The response[0] contains a response code. 
        # A 250 response code means the request was successful
        # A 550 response code means the user does not exist
        if response[0] == 250:
            print(f"{email_address} exists.")
            return True
        elif response[0] == 550:
            print(f"{email_address} does not exist.")
            return False
        else:
            print(f"Received {response[0]} response.")
            return False
    except (smtplib.SMTPConnectError, socket.timeout) as e:
        print(f"Error: {e}")
        return False

def check_email(email_address):
    # Split the email address at '@' and get the domain
    domain = email_address.split('@')[1]
    return validate_email(email_address, domain)

# test the code with an example email
check_email("user0.user0.user0@gmail.com")
