import win32com.client as win32
import psutil
import os
import subprocess
import pandas as pd
import time




def enviar_email():
    leitura = pd.read_excel(r"C:\Users\Deyvid\Desktop\Controle Financeiro\ControleFinanceiro.xlsx")
    total = leitura['Unnamed: 8'][11]


    def send_notification():
        t = time.localtime()
        if t[1]>0 and t[1]<9:
            m = f'0{t[1]}'

        else:
            m = t[1]


        if m == '01' :
            mes = 'Janeiro'
        elif m == '02':
            mes = 'Fevereiro'
        elif m == '03':
            mes = 'Março'
        elif m == '04':
            mes = 'Abril'
        elif m == '05':
            mes = 'Maio'
        if m == '06' :
            mes = 'Junho'
        elif t[1] == '07':
            mes = 'Julho'
        elif m == '08':
            mes = 'Agosto'
        if m == '09' :
            mes = 'Setembro'
        elif m == '10':
            mes = 'Outubro'
        elif m == '11':
            mes = 'Novembro'
        elif m == '12':
            mes = 'Dezembro'
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'deyvinholins18@gmail.com; angelicasantos444@gmail.com'
        mail.Subject = 'Relatório de controle financeiro referente ao mês de Abril'
        mail.HTMLbody = f'''
        <p>Segue o total de gasto referente ao mês de <strong>{mes.capitalize()}</strong></p>
            
        <p>O seu total de gasto neste mês foi:<strong>{total}</strong></p>   
       

        <p>EM anexo, contém mais informações detalhadas</p><br><br><br><br><br><br><br><br><br>
        
        
        <p> &#9888; &#9888; &#9888; <strong>ATENÇÂO!!!</strong> Tente não gastar muito para dormir mais tranquilo(a) e ficar bem financeiramente.</p>

        <p>Arquivo modificado na data: <strong>{t[2]}/{m}/{t[0]} às {t[3]}:{t[4]}:{t[5]}</strong></p>'''

        anexo = r"C:\Users\Deyvid\Desktop\Controle Financeiro\ControleFinanceiro.xlsx"
        mail.Attachments.Add(anexo)
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

enviar_email()