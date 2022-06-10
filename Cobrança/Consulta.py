import sched
import yagmail
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import datetime
import pandas as pd
import warnings
warnings.simplefilter("ignore")

scheduler = sched.scheduler()
def consulta():
    df = pd.read_excel(r'C:/Users/powerbi/Google Drive/IMPORTAÇÃO/FINANCEIRO/CONTAS_RECEBER_VENCIDO.xlsx', sheet_name='LOJAS', engine='openpyxl')

    df = pd.DataFrame(df, columns=['CPFCNPJ', 'NOMEPARCEIRO', 'CODPORTADOR','CODEMPSW', 'NOMEVENDEDOR', 'VLRTITULO', 'DTVENC', 'TRACKING'])

    df['VLRTITULO'] = df['VLRTITULO'].map(lambda x: float(x))
    #Identificando os status
    df.loc[df['CODPORTADOR'] == 991, 'STATUS'] = 'CHEQUE'
    df.loc[df['CODPORTADOR'] == 1000, 'STATUS'] = 'CHEQUE'
    df.loc[df['CODPORTADOR'] == 998, 'STATUS'] = 'A EXECUTAR'
    df.loc[df['CODPORTADOR'] == 810, 'STATUS'] = 'COBRANCA SOLUTE'
    df.loc[df['CODPORTADOR'] == 345, 'STATUS'] = 'COBRANCA SOLUTE'
    df.loc[df['CODPORTADOR'] == 800, 'STATUS'] = 'COBRANCA LITIGIO'
    df.loc[df['CODPORTADOR'] == 343, 'STATUS'] = 'COBRANCA LITIGIO'
    df.loc[df['CODPORTADOR'] == 808, 'STATUS'] = 'COBRANCA MAC'
    df.loc[df['CODPORTADOR'] == 809, 'STATUS'] = 'DEVOLUCAO EMPRESA COBRANCA'
    df.loc[df['CODPORTADOR'] == 802, 'STATUS'] = 'DEVOLUCAO EMPRESA COBRANCA'
    df.loc[df['CODPORTADOR'] == 801, 'STATUS'] = 'DEVOLUCAO EMPRESA COBRANCA'
    df.loc[df['CODPORTADOR'] == 811, 'STATUS'] = 'DEVOLUCAO EMPRESA COBRANCA'
    df.loc[df['CODPORTADOR'] == 804, 'STATUS'] = 'CLIENTE POR PERDA'
    df.loc[df['CODPORTADOR'] == 803, 'STATUS'] = 'CLIENTE EXECUCAO'
    df.loc[df['CODPORTADOR'] == 757, 'STATUS'] = 'EX FUNCIOARIO'
    df.loc[df['CODPORTADOR'] == 758, 'STATUS'] = 'FALECIDO'
    df.loc[df['CODPORTADOR'] == 806, 'STATUS'] = 'DIRETORIA'
    df.loc[df['CODPORTADOR'] == 805, 'STATUS'] = 'CLIENTE RJ'
    df.loc[df['CODPORTADOR'] == 0, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 1, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 2, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 4, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 33, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 225, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 341, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 342, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 344, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 353, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 354, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 360, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 422, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 990, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 995, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 997, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 999, 'STATUS'] = 'COBRANCA DF'
    df.loc[df['CODPORTADOR'] == 993, 'STATUS'] = 'CARTAO'
    df.loc[df['CODPORTADOR'] == 994, 'STATUS'] = 'CARTAO'

    #CLIENTES DIRETORIA
    df.loc[df['CPFCNPJ'] == '16.622.166/0005-03', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '40.882.060/0001-08', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '810.978.404-68', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '12.275.715/0001-36', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '58.317.751/0014-30', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.319.853/0001-44', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '07.196.033/0054-00', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.291.177/0001-48', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '03.965.584/0015-23', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '30.734.711/0001-50', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '430.863.904-25', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '07.196.033/0041-95', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '16.622.166/0008-56', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '07.523.790/0001-39', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '03.965.584/0013-61', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '30.817.641/0001-02', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '03.965.584/0019-57', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '18.650.667/0001-03', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '32.202.670/0001-87', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '15.350.602/0014-60', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '12.217.832/0002-24', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '29.067.113/0345-03', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '12.854.865/0001-02', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '08.701.062/0001-32', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '03.237.583/0045-88', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '24.360.910/0001-43', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '07.196.033/0055-90', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '73.410.326/0119-52', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '28.142.800/0014-80', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.110.989/0001-40', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '02.536.066/0015-21', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '08.944.084/0001-23', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '05.586.713/0001-00', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '30.248.954/0001-89', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '869.620.864-15', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '07.451.885/0006-07', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.785.202/0001-40', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.970.887/0037-05', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.143.246/0001-76', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '12.200.259/0001-65', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '13.399.232/0003-78', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '10.091.510/0001-75', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '07.196.033/0056-71', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '01.340.982/0001-23', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '11.049.806/0001-90', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '09.067.562/0001-27', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '05.659.785/0010-13', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '29.067.113/0364-68', 'STATUS'] = 'DIRETORIA'
    df.loc[df['CPFCNPJ'] == '01.596.018/0001-60', 'STATUS'] = 'DIRETORIA'
    # Dia de hoje
    data_atual = datetime.datetime.now()
    data_em_texto = data_atual.strftime('%m/%d/%Y')
    data = datetime.datetime.strptime(data_em_texto, '%m/%d/%Y')
    # Coluna Quantidade de Dias
    df['quantidade_dias'] = abs((data - df['DTVENC'].astype('datetime64[ms]')))
    # Transformanfo o para inteiro e retirando o nome "dias"
    df['quantidade_dias'] = df['quantidade_dias'].astype('timedelta64[D]').astype(int)
    # Localizando os rotativos
    df.loc[df['quantidade_dias'] < 40, 'ROTATIVO'] = 'ROTATIVO'
    # Inserindo o Rotativo no Status
    df.loc[df['ROTATIVO'] == 'ROTATIVO', 'STATUS'] = 'ROTATIVO 40 DIAS' 
    # Criando Ano Convertendo para texto
    df['ANO'] = df['DTVENC'].apply(lambda x: str(x)[:4])
    # Mês
    df['MES'] = df['DTVENC'].dt.month
    df.loc[df['MES'] == 1, 'MES'] = 'Janeiro'
    df.loc[df['MES'] == 2, 'MES'] = 'Fevereiro'
    df.loc[df['MES'] == 3, 'MES'] = 'Março'
    df.loc[df['MES'] == 4, 'MES'] = 'Abril'
    df.loc[df['MES'] == 5, 'MES'] = 'Maio'
    df.loc[df['MES'] == 6, 'MES'] = 'Junho'
    df.loc[df['MES'] == 7, 'MES'] = 'Julho'
    df.loc[df['MES'] == 8, 'MES'] = 'Agosto'
    df.loc[df['MES'] == 9, 'MES'] = 'Setembro'
    df.loc[df['MES'] == 10, 'MES'] = 'Outubro'
    df.loc[df['MES'] == 11, 'MES'] = 'Novembro'
    df.loc[df['MES'] == 12, 'MES'] = 'Dezembro'
    # Complemento STATUS
    df.loc[(df['STATUS'] == 'COBRANCA DF') & (df['ANO'] == '2018'), 'STATUS'] = 'COBRANCA DF 2018'
    df.loc[(df['STATUS'] == 'COBRANCA DF') & (df['ANO'] == '2019'), 'STATUS'] = 'COBRANCA DF 2019'
    df.loc[(df['STATUS'] == 'COBRANCA DF') & (df['ANO'] == '2020'), 'STATUS'] = 'COBRANCA DF 2020'
    df.loc[(df['STATUS'] == 'COBRANCA DF') & (df['ANO'] == '2021'), 'STATUS'] = 'COBRANCA DF 2021'
    df.loc[(df['STATUS'] == 'COBRANCA DF') & (df['ANO'] == '2022'), 'STATUS'] = 'COBRANCA DF 2022'

    df.loc[df['ANO'] == '2015', 'ANO2'] = '2015 a 2017'
    df.loc[df['ANO'] == '2016', 'ANO2'] = '2015 a 2017'
    df.loc[df['ANO'] == '2017', 'ANO2'] = '2015 a 2017'
    df.loc[(df['STATUS'] == 'A EXECUTAR') & (df['ANO2'] == '2015 a 2017'), 'STATUS'] = 'A EXECUTAR (15 A 17)'
    df.loc[(df['STATUS'] == 'CLIENTE RJ') & (df['ANO2'] == '2015 a 2017'), 'STATUS'] = 'CLIENTE RJ (15 A 17)'
    df.loc[(df['STATUS'] == 'CLIENTE EXECUCAO') & (df['ANO2'] == '2015 a 2017'), 'STATUS'] = 'CLIENTE EXECUCAO (15 A 17)'

    df['VLRTITULO'] = df['VLRTITULO'].round(2)
    df['VLRTITULO'] = df['VLRTITULO'].apply(lambda x: str(x).replace('.',','))

    df = pd.DataFrame(df, columns=['CPFCNPJ', 'NOMEPARCEIRO', 'STATUS','CODEMPSW', 'NOMEVENDEDOR', 'VLRTITULO', 'DTVENC', 'ANO', 'TRACKING'])

    df.to_csv('consulta.csv', sep=';', index=False)
    # Imprime o Log
    agora = datetime.datetime.now()
    print(f'Atualizado com sucesso! Horário: {agora}')
    # Agendamento do job
    scheduler.enter(delay=604800, priority=0, action=consulta)

consulta()
scheduler.run(blocking=True)