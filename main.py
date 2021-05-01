import win32com.client as win32
import psutil
import os
import subprocess
import pandas as pd
import time
import openpyxl
from datetime import datetime


wb = openpyxl.load_workbook(r"C:\Users\Deyvid\Desktop\ControleFinanceiro.xlsx")

leitura = pd.read_excel(r"C:\Users\Deyvid\Desktop\ControleFinanceiro.xlsx", sheet_name=f"{wb.active.title}")
total = leitura['Unnamed: 8'][11]

titulo = leitura['Unnamed: 0'][0]
de = leitura['Unnamed: 0'][2:30]
va = leitura['Unnamed: 1'][2:30]
dt_venc = leitura['Unnamed: 2'][2:30]
dt_pag = leitura['Unnamed: 3'][2:30]
v_pag = leitura['Unnamed: 4'][2:30]
deve = leitura['Unnamed: 5'][2:30]
sts = leitura['Unnamed: 6'][2:30]

lista_desc = []
lista_valor = []
lista_data_venc = []
lista_data_pag = []
lista_valor_pag = []
lista_dev = []
lista_sts = []


for d in de:
    lista_desc.append(d)

for v in va:
    lista_valor.append(v)

for dt_v in dt_venc:
    lista_data_venc.append(dt_v)

for dt_p in dt_pag:
    lista_data_venc.append(dt_p)

for vl_pag in v_pag:
    lista_valor_pag.append(vl_pag)

for dev in deve:
    lista_dev.append(dev)

for s in sts:
    lista_sts.append(s)


def remove_nan():
    global  list_descricao,list_valor,list_data_vencimento,list_data_pagamento,list_valor_pago,list_devendo,list_status

    list_descricao = []
    for d in lista_desc:
        descricao = str(d)
        if descricao != 'nan':
            list_descricao.append(descricao)

    list_valor = []
    for v in lista_valor:
        valor = str(v)
        if valor != 'nan':
            list_valor.append(valor)

    list_data_vencimento = []
    for d in lista_data_venc:
        da = str(d)
        if da != 'nan':
            dat = da[0:10]
            date = datetime.strptime(dat, '%Y-%m-%d').date()
            data = date.strftime('%d/%m/%Y')

            list_data_vencimento.append(data)

    list_data_pagamento = []
    for d in lista_data_pag:
        da = str(d)
        if da != 'nan':
            dat = da[0:10]
            date = datetime.strptime(dat, '%Y-%m-%d').date()
            data = date.strftime('%d/%m/%Y')
            list_data_pagamento.append(data[0:10])

    list_valor_pago = []
    for v in lista_valor_pag:
        valor_pago = str(v)
        if valor_pago != 'nan':
            list_valor_pago.append(valor_pago)

    list_devendo = []
    for d in lista_dev:
        devendo = str(d)
        if devendo != 'nan':
            list_devendo.append(devendo)

    list_status = []
    for s in lista_sts:
        status = str(s)
        if status != 'nan':
            list_status.append(status)

remove_nan()

def send_notification():
    t = time.localtime()
    if t[1] > 0 and t[1] < 9:
        m = f'0{t[1]}'
    else:
        m = t[1]

    if t[2] > 0 and t[2] < 9:
        d = f'0{t[2]}'
    else:
        d = t[2]

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'deyvinholins18@gmail.com; angelicasantos444@gmail.com'
    mail.Subject = f'Relatório de controle financeiro referente ao mês de {wb.active.title}'
    mail.HTMLbody = f'''
    <p>Segue o total de gasto referente ao mês de <strong>{wb.active.title.capitalize()}</strong></p>

    <p>O seu total de gasto neste mês foi:<strong>{total}</strong></p>        
    

    <p>EM anexo, contém mais informações detalhadas</p><br> 
    
    <p> &#9888; &#9888; &#9888; <strong>ATENÇÂO!!!</strong> Tente não gastar muito para dormir mais tranquilo(a) e ficar bem financeiramente.</p>
    <p>Arquivo modificado na data: <strong>{d}/{m}/{t[0]} às {t[3]}:{t[4]}:{t[5]}</strong></p>'''

    anexo = r"C:\Users\Deyvid\Desktop\ControleFinanceiro.xlsx"
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


