# Send message via gmail
# Input: gmail account, user and password in config.py
# 
# 07.02.2023
# v1
# Kadi Jairus

import smtplib, ssl, os
import config


# https://mljar.com/blog/python-send-email/ 2023-02-07
def mailsender(receiver_email, subject, message):
    try:
        sender_email = config.user
        apppassword = config.password
        
        if sender_email is None or apppassword is None:
            # no email address or password
            # something is not configured properly
            print("Did you set email address and password correctly?")
            return False
        
        port = 465  # For SSL
        smtp_server = "smtp.gmail.com"
        
        message = f"""\
Subject: {subject}
        
{message}
        
Tervitustega
Meilirobot

"""

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, apppassword)
            server.sendmail(sender_email, receiver_email, message)
            print("Mailsender OK!")
        return True

    except Exception as e:
        print(sender_email)
        print(receiver_email)
        print(message)
        print("Error!")
        print(str(e))
    return False

message1 = "Põhitekst asub siin"
message1 = message1.encode('utf-8').decode('latin-1')
print(message1)
paus = 1/0
mailsender("kadijairus@gmail.com","Teavitus",message1)