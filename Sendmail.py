# Set up sceduled task, check Excel and send email if changes are detected
# Input: new and archive version of the same file; gmail account; user and password in config.py; Task Scheduler
# Output: email
# 07.02.2023
# v1
# Kadi Jairus

import pandas as pd
import shutil
import smtplib, ssl, os
#config.py in the same folder: user = "***" password = "b***"
import config


#dir_old = r'C:\Intel\KadiJairus\Arhiiv\2023 Ajakava.xlsx'
#dir_new = r'\\srvlaste\Yhendlabor\GE_Geneetikakeskus\Puhkused_koolitused\2023 Ajakava.xlsx'
# Asukohad testkaustas
dir_old = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\Sendmail_arhiiv\2023 Ajakava.xlsx'
dir_new = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\2023 Ajakava.xlsx'

def sheetcomparer(dir_old,dir_new,sheetname):
    df_old = pd.read_excel(dir_old,sheet_name=sheetname,header=1,usecols=["Nimi","Algus","Lõpp"])
    df_new = pd.read_excel(dir_new,sheet_name=sheetname,header=1,usecols=["Nimi","Algus","Lõpp"])
    df = df_new.compare(df_old, align_axis=0, keep_shape = False).rename(index={'self': 'uus', 'other': 'vana'},level=-1)
    df["Leht"] = sheetname
#    df = df[['Nimi', 'self'],['Algus', 'self'],['Lõpp', 'self'],['Leht', '']]
    df.replace({'NaN': '', 'NaT': ''})
    print(df)

sheetcomparer(dir_old,dir_new,"P")


teesiinpaus = 1/0
shutil.copy(dir_new,dir_old)
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
            server.sendmail(sender_email, receiver_email, message.encode('utf-8'))
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
message1 = message1.encode('utf-8').decode('utf-8')
#print(message1)
#paus = 1/0
mailsender("kadijairus@gmail.com","Teavitus",message1)
