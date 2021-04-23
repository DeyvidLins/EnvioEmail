import win32com.client as win32
import psutil
import os
import subprocess
import pandas as pd

def enviar_email():
    leitura = pd.read_excel(r"C:\Users\Deyvid\Desktop\Controle Financeiro\ControleFinanceiro.xlsx")
    total = leitura['Unnamed: 8'][11]
    print(total)



    def send_notification():
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'deyvinholins18@gmail.com; angelicasantos444@gmail.com'
        mail.Subject = 'Relatório de Controle Financeiro Referente ao mês de Abril'
        mail.body = f'''
        Segue o total gasto referente ao mês de Abril
        total:{total}'''
        mail.Send()


    # Open Outlook.exe. Path may vary according to system config
    # Please check the path to .exe file and update below

    def open_outlook():
        try:
            subprocess.call([r'"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"'])
            os.system(r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE")

        except:
            print("Outlook não abriu com sucesso")

    # Checking if outlook is already opened. If not, open Outlook.exe and send email
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            flag = 1
            break
        else:
            flag = 0

    if (flag == 1):
        print("E-mail enviado com Sucesso!")
        send_notification()
    else:
        open_outlook()
        send_notification()


