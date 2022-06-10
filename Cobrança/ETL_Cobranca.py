from lib2to3.pgen2.pgen import DFAState
import mysql.connector
import pandas as pd
import sched
import datetime
import warnings
warnings.simplefilter("ignore")


scheduler = sched.scheduler()
def envio():
    # Conecta o Banco
    mydb = mysql.connector.connect(host="10.0.1.251", user="cobranca", password="Dfp@321", database="cobranca")  
    mycursor = mydb.cursor()
    # Realiza o Select
    mycursor.execute("SELECT * FROM tbcobranca")
    myresult = mycursor.fetchall()
    # Transforma o Select em um DataFrame
    df = pd.DataFrame(myresult, columns = ['CLIENTE', 'PORTADOR', 'VALOR', 'COBRADORAS', 'MES', 'LOJA', 'DICIONARIO', 'TRACKING', 'SEMANA', 'ID'])
    df = df.drop_duplicates()
    # Salva o DataFrame em Excel
    df.to_excel("ETL_Base.xlsx", index = False, sheet_name='etl')
    # Imprime o Log
    agora = datetime.datetime.now()
    print(f'ETL da cobrança atualizado com sucesso! Horário: {agora}')
    # Agendamento do job
    scheduler.enter(delay=5, priority=0, action=envio)

envio()
scheduler.run(blocking=True)