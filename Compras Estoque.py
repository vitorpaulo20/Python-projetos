#!/usr/bin/env python
# coding: utf-8

import sched
import yagmail
import datetime
import pandas as pd
import openpyxl
import warnings
warnings.simplefilter("ignore")

# 86400 segundos é um dia
#compras@dafontepneus.com.br

# ### Envio da planilha de estoque e consumo para pedidos Compras

scheduler = sched.scheduler()
def printa():
    df = pd.read_excel(r'C:/Users/powerbi/Google Drive/IMPORTAÇÃO/ESTOQUE/ESTOQUE.xlsx', engine='openpyxl')
    df.to_excel('ESTOQUE.xlsx')

    ds = pd.read_excel(r'C:/Users/powerbi/Google Drive/IMPORTAÇÃO/PRODUCAO/CONSUMO_PRODUCAO.xlsx', engine='openpyxl')
    ds.to_excel('CONSUMO_PRODUCAO.xlsx')
    user = yagmail.SMTP(user='corporativodfp@gmail.com', password='xthvzszwigmbfndf')
    user.send(to='compras@dafontepneus.com.br', subject='Base estoque e consumo', contents='Olá!,\n \nSeguem as planilhas base para Soap planilha de pedidos, segue ordem para alimentar planilha: \n 1° CONSUMO_PRODUCAO.xlsx \n 2° ESTOQUE.xlsx  .\nAtt.,\nP.A.U.L.O.', attachments=['ESTOQUE.xlsx','CONSUMO_PRODUCAO.xlsx'])
    agora = datetime.datetime.now()
    print(f'Estoque e consumo enviados com sucesso ! Horário: {agora}')
    
    scheduler.enter(delay=76000, priority=0, action=printa)
    
printa()
scheduler.run(blocking=True)
