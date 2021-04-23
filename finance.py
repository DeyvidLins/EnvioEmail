import os
import time
from main import enviar_email
import subprocess

os.startfile("ControleFinanceiro.xlsx")

i = 0

while i <= 1:
    time.sleep(10)
    try:
        a = open("ControleFinanceiro.xlsx", "r+")
        print("Arquivo está fechado")
        a.close()

        enviar_email()
        break
    except IOError:
        print("Arquivo está aberto")


