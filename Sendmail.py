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
#dir_original = = r'\\srvlaste\Yhendlabor\GE_Geneetikakeskus\Puhkused_koolitused\2023 Ajakava.xlsx'
#dir_old = r'C:\Intel\KadiJairus\Arhiiv\2023 Ajakava uus.xlsx'
# Asukohad testkaustas
dir_old = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\Sendmail_arhiiv\2023 Ajakava.xlsx'
dir_original = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\2023 Ajakava.xlsx'
dir_new = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\Sendmail_arhiiv\2023 Ajakava uus.xlsx'

try:
    shutil.copy(dir_original,dir_new)
except:
    print("Jätkan")






allsheetnames = ["P","T"]

names = pd.read_excel(dir_new,sheet_name='Töötajad',header=2,usecols=[4,18])
names.columns = ['Name', 'Dpt']
names.dropna()
names = names[names['Dpt'].str.contains("kliiniline geneetika")==True]
print(names)
paus =1/0

names.drop(names.index[29:101], inplace=True)
pd.set_option('display.max_rows', None)
names.drop(names.index[118:142], inplace=True)
print(names)
names.drop(names.index[143:144], inplace=True)
#names = names.drop(names.loc[30:102].index, inplace=True)
#names = names.drop(index[(names[30:102]) & (names[120:152])].index)
names = names['Töötajad'].unique()
print(names)
names = list(names["Töötajad"])
print(names)


teesiinpaus = 1/0

def sheetcomparer(dir_old,dir_new,sheetname):
    df_old = pd.read_excel(dir_old,sheet_name=sheetname,header=1,usecols=["Nimi","Algus","Lõpp"])
    df_new = pd.read_excel(dir_new,sheet_name=sheetname,header=1,usecols=["Nimi","Algus","Lõpp"])
    df = df_new.compare(df_old, align_axis=0, keep_shape = False).rename(index={'self': 'uus', 'other': 'enne muutmist'},level=-1)
    df["Leht"] = sheetname
    df.fillna('-')
    return(df)


try:
    df_appended=[]
    if len(allsheetnames) > 0:
        for asheetname in allsheetnames:
            print(asheetname)
            df_sheet = sheetcomparer(dir_old,dir_new,asheetname)
            df_appended.append(df_sheet)
            print(df_appended)
        df = pd.concat(df_appended)
        df = df.fillna('-')
        print(df)
    else:
        print("No sheetnames!")
except Exception as e:
    print("Sheetcomparer error!")
    print(str(e))

nimedefilter = (df['Nimi'].isin(names))
df = df.loc[nimedefilter]
print("FILTREERITUD")
#Kuupäevalt kellaaeg maha, sidekriips vahele, "Leht" maha
print(df)

#teesiinpaus = 1/0

#shutil.copy(dir_new,dir_old)

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
Subject: Ajakavas {subject} muudatust

Tere!

Ajakavas on sellised muudatused:

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

#message1 = "Põhitekst asub siin"
#df = df.encode('utf-8').decode('utf-8')
#print(message1)
#paus = 1/0
mailsender("kadijairus@gmail.com", len(df),df)

os.remove(dir_old)
os.rename(dir_new,dir_old)

print("OK!")
