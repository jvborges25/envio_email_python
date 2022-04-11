from concurrent.futures import thread
from unicodedata import name
from urllib import response
from wsgiref import headers
import pandas as pd
import datetime
import smtplib
import time
import requests
from win10toast import ToastNotifier

GMAIL_ID = 'seu_email'
GMAIL_PWD = 'sua_senha'

toast = ToastNotifier()

def sendEmail(to, sub, msg):

    gmail_obj = smtplib.SMTP('smtp.gmail.com', 587)

    gmail_obj.starttls()

    gmail_obj.login(GMAIL_ID, GMAIL_PWD)

    gmail_obj.sendmail(GMAIL_ID, to, f"subject : {sub}\n\n{msg}")

    gmail_obj.quit()

    print("email enviado para" + str(to) + "assunto" + str(sub) + "mensagem :" + str(msg))

    toast.show_toast("Email enviado!", f"{name} e-mail enviado com sucesso", threaded = True, 
        icon_path =None, duracao = 6)

    while toast.notification_active():
        time.sleep(0.1)


def sendsms(to, msg, name, sub):

    url = "http://www.fast2sms.com"
    payload = f"sender_id=FSTSMS&message={msg}&language=pt&route=p&numbers={to}"

    headers = {
        'authorization' : "API_KEY_HERE",
        'Content-Type' :"application/x-www-form-urlencoded",
        'Cache-Control' : "no-cache",
    }

    response_obj = requests.request("POST", url, data = payload, headers=headers)

    print(response_obj.text)
    print("SMS enviado para" + str(to) + "Assunto :" + str(sub) + "Mensagem :" + str(msg))

    toast.show_toast("SMS enviado!", f"{name} Mensagem Enviada com sucesso!", threaded = True,
        icon_path = None, duration = 6)


    while toast.notification_Active():
        time.sleep(0.1)

if __name__=="__main__":

    dataframe = pd.read_excel("arquivo_exel.xlsx")

    today = datetime.datetime.now().strftime("%D-%M")

    yearNow = datetime.datetime.now().strftime("%Y")

    writeInd = []

    for index, item in dataframe.interrows():
        msg = "To passando para te desejar tudo de bom e que esse dia possa se repetir muitas e muitas vezes " + str(item['NAME'])

        bday = item['Parabens'].strftime("%D-%M") 

            if (bday == today) and yearNow not in str(item['Year']):
                sendEmail(item['Email'], "Feliz Aniversario", msg) + sendsms(item['Contact'], msg, item['name'], "Feliz Aniversario")
                writeInd.append(index) 

        
    for i in writeInd:
        yr = dataframe.loc[i, 'Year']

        dataframe.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)

    dataframe.to_excel('exelsheet.xlsx', index = False)



