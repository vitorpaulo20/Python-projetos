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
#hermileide.bezerra@dafontepneus.com.br

# ### Envio da planilha de faturamento diário para Hermileide

scheduler = sched.scheduler()
def printa():
    df = pd.read_excel(r'C:/Users/powerbi/Google Drive/IMPORTAÇÃO/VENDAS/VENDAS_GRUPO.xlsx', engine='openpyxl')
    df.to_csv('VENDAS_GRUPO.csv')
    user = yagmail.SMTP(user='corporativodfp@gmail.com', password='xthvzszwigmbfndf')
    user.send(to='hermileide.bezerra@dafontepneus.com.br', subject='Faturamento diário', contents='Olá!,\n \nSegue a Planilha de faturamento.\n \n \nAtt.,\nJ.A.R.V.I.S.', attachments='C:/Users/powerbi/Google Drive/IMPORTAÇÃO/VENDAS/VENDAS_GRUPO.csv')
    agora = datetime.datetime.now()
    print(f'Faturamento enviado com sucesso! Horário: {agora}')
    scheduler.enter(delay=76000, priority=0, action=printa)
    
printa()
scheduler.run(blocking=True)