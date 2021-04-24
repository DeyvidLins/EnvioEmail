import os
import time


os.startfile(r"C:\Users\Deyvid\Desktop\ControleFinanceiro.xlsx")

i = 0

while i <= 1:
    time.sleep(10)
    try:
        a = open(r"C:\Users\Deyvid\Desktop\ControleFinanceiro.xlsx", "r+")
        print("Arquivo está fechado")
        a.close()
        os.system(r'python "C:\Users\Deyvid\Desktop\Controle Financeiro\main.py" ')

        break
    except IOError:
        print("Arquivo está aberto")


