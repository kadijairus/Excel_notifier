# Set up sceduled task, check Excel and send email if changes are detected
# Input: new and archive version of the same file; gmail account; user and password in config.py; Task Scheduler
# Output: email
# 13.02.2023
# v2
# Kadi Jairus


import datetime
import win32com.client as win32
import pandas as pd
import shutil
import smtplib, ssl, os
#config.py in the same folder: user = "***" password = "b***"
import config


# TheoretiCAL https://stackoverflow.com/questions/6332577/send-outlook-email-via-python
def outlookmailsender(receiver_email, subject, message,lastchanged):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = receiver_email
    mail.Subject = 'Ajakavas ' + subject
    mail.Body = f"""\
Tere!

Ajakava tabelis {subject}.

{message}

Tabelit muudeti viimati {lastchanged}.
        
Tervitustega
Meilirobot

"""
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    mail.Send()


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
Subject: Ajakavas {subject}

Tere!

Ajakavas {subject}.
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

#mailsender("kadijairus@gmail.com", "Esimene def laetud","Esimene def laetud")

dir_old = r'C:\Intel\KadiJairus\2023 Ajakava.xlsx'
dir_original = r'\\srvlaste\Yhendlabor\GE_Geneetikakeskus\Puhkused_koolitused\2023 Ajakava.xlsx'
dir_new = r'C:\Intel\KadiJairus\2023 Ajakava uus.xlsx'
# Asukohad testkaustas
#dir_old = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\Sendmail_arhiiv\2023 Ajakava.xlsx'
#dir_original = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\2023 Ajakava.xlsx'
#dir_new = r'D:\Users\loom\Desktop\Pisi\T88\Python jms\Sendmail_arhiiv\2023 Ajakava uus.xlsx'

#mailsender("kadijairus@gmail.com", "Tabelite asukohad laetud","Tabelite asukohad laetud")

try:
    shutil.copy(dir_original,dir_new)
except:
    print("Jätkan")

try:
    oslastchanged = os.path.getmtime(dir_original)
    lastchanged = datetime.datetime.fromtimestamp(oslastchanged).strftime("%d.%m.%Y %H.%M")
except:
    lastchanged = 'teadmata ajal'
    
#mailsender("kadijairus@gmail.com", "Üritatud uuendada","Üritatud uuendada dir-ist")

allsheetnames = ["P","T","H","K","E","M"]

names = pd.read_excel(dir_new,sheet_name='Töötajad',header=2,usecols=[4,18])
names.columns = ['Nimi', 'Dpt']
names.dropna()
names = names[names['Dpt'].str.contains("kliiniline geneetika")==True]
names = names['Nimi']
#names = list(names["Name"])
print(names)

#mailsender("kadijairus@gmail.com", "Nimed laetud","Nimed laetud")


def sheetcomparer(dir_old,dir_new,sheetname):
    df_old = pd.read_excel(dir_old,sheet_name=sheetname,header=1,usecols=["Nimi","Algus","Lõpp"])
    df_new = pd.read_excel(dir_new,sheet_name=sheetname,header=1,usecols=["Nimi","Algus","Lõpp"])
#    nimedefilter = (df_new['Nimi'].isin(names))
#    df_new = df_new.loc[nimedefilter]
    df = df_new.compare(df_old, align_axis=0, keep_shape = False, keep_equal = True).rename(index={'self': 'lisatud ajad', 'other': 'enne muutmist'},level=-1)
    df["Leht"] = sheetname
    df.fillna('-')
    return(df)

df = []
try:
    df_appended=[]
    if len(allsheetnames) > 0:
        for asheetname in allsheetnames:
            print(asheetname)
            df_sheet = sheetcomparer(dir_old,dir_new,asheetname)
            df_appended.append(df_sheet)
            print(df_appended)
        df = pd.concat(df_appended)
        df['Algus'] = pd.to_datetime(df['Algus'])
        df['Algus'] = df['Algus'].dt.strftime('%d.%m.%Y')
        df['Lõpp'] = pd.to_datetime(df['Lõpp'])
        df['Lõpp'] = df['Lõpp'].dt.strftime('%d.%m.%Y')
        print('Enne filtreerimist')
        print(df)
#        df['Name'] =df['Nimi']
# Filtrit ei õnnestunud ettepoole panna. Näitab ainult kuupäeva muutmisel tühja lahtrit
        nimedefilter = (df['Nimi'].isin(names))
        df = df.loc[nimedefilter]
        df.fillna('-')
#        df = df.to_string(index=False)
        df = df[['Leht','Algus','Lõpp','Nimi']]
        print("FILTREERITUD")
        print(df)
    else:
        print("No sheetnames!")
except Exception as e:
    print("Sheetcomparer error!")
    print(str(e))
    raise SystemExit

#mailsender("kadijairus@gmail.com", "Df tehtud","Df tehtud")


if len(df) > 0:
    subject = "on tehtud muudatusi"
else:
    subject = "ei ole ühtegi muudatust"
    df = ''
#print(df)

#teesiinpaus = 1/0

#shutil.copy(dir_new,dir_old)
    
#mailsender("kadijairus@gmail.com", "len Df tehtud","len Df tehtud")



#message1 = "Põhitekst asub siin"
#df = df.encode('utf-8').decode('utf-8')
#print(message1)
#paus = 1/0
outlookmailsender("kadi.jairus@kliinikum.ee", subject,df,lastchanged)
outlookmailsender("triin.tago@kliinikum.ee", subject,df,lastchanged)
#outlookmailsender("katlin.kraavik@kliinikum.ee", subject,df,lastchanged)
#outlookmailsender("kadijairus@gmail.com", subject,df,lastchanged)
#mailsender("kadijairus@gmail.com", 'Teate koopia',df)

os.remove(dir_old)
os.rename(dir_new,dir_old)

print("OK!")
