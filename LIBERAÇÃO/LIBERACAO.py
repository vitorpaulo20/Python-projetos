#!/usr/bin/env python
# coding: utf-8

import sched
import yagmail
import datetime
import pandas as pd
import openpyxl
import re
from sys import exit
import pyautogui as pag
import warnings
from openpyxl import Workbook, load_workbook
warnings.simplefilter("ignore")

# 86400 segundos é um dia




def clear():
    try:
        import os
        lines = os.get_terminal_size().lines
    except AttributeError:
        lines = 130
    print("\n" * lines)



scheduler = sched.scheduler()

def printa():
    clear()
    planilha = load_workbook("LIBERACAO.XLSX")
    aba_ativa = planilha.active
    
    df = pd.read_excel(r'LIBERACAO.xlsx', engine='openpyxl')
    
    user = yagmail.SMTP(user='corporativodfp@gmail.com', password='xthvzszwigmbfndf')
    
    
    
    loja = input("Digite a loja para solicitar liberação: ")
    t1 = re.compile(r'^\d{2}$')
    checarLOJA = t1.findall(loja) 
    if not checarLOJA:
      print()
      scheduler.enter(delay=100, priority=0, action=pag.alert(text=f'{loja} Não é uma loja, por favor inserir dois digitos exemplo: 01  se for loja 01', title="Loja errada !!"))
      printa()
    else:   
      print('LOJA INSERIDA COM SUCESSO !!')   
   
        
        
    cnpjcpf = input("Digite o numero do cnpj ou cpf: ")
    t2 = re.compile(r"(^[0-9]{2}[\.]?[0-9]{3}[\.]?[0-9]{3}[\/]?[0-9]{4}[-]?[0-9]{2})|([0-9]{3}[\.]?[0-9]{3}[\.]?[0-9]{3}[-]?[0-9]{2}$)")
    checarCNCP = t2.findall(cnpjcpf) 
    if not checarCNCP:
      print()
      scheduler.enter(delay=100, priority=0, action=pag.alert(text=f'{cnpjcpf} Escreva sem caracteres especiais apenas numeros 11 digitos para cpf e 14 para cnpj', title="INVALIDO !!!"))
      printa()
    else:   
      print('LOJA INSERIDA COM SUCESSO !!') 
    
    pv = input("Digite o numero da PV: ")
        
    aba_ativa["a2"]=cnpjcpf
    aba_ativa["b2"]=pv
    
    
   
    
   
    
   
    
    planilha.save('LIBERACAO.xlsx')
    
    user.send(to='paulo.almeida@dafontepneus.com.br', subject=f'{loja}', contents=f'{pv}', attachments=['LIBERACAO.xlsx'])
    agora = datetime.datetime.now()
    print(f'Liberação enviada com sucesso! Horário: {agora} \n\n')
    
    scheduler.enter(delay=10, priority=0, action=printa)    


printa()   
scheduler.run(blocking=True)


