# Send message via gmail
# Input: gmail account, user and password in config.py
# 
# 07.02.2023
# v1
# Kadi Jairus

import smtplib, ssl, os
import config


port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = config.user
apppassword = config.password
receiver_email = "kadijairus@gmail.com"
variable = 5

message = f"""\
Subject: Ajakava tabelis {variable} muudatust

This message is sent from Python.

Tervitustega

Meilirobot

"""

context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    server.login(sender_email, apppassword)
    server.sendmail(sender_email, receiver_email, message)
    
print("Tehtud")
