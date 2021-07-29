import numpy as np                                  # Usada para trabalhar com números e matrizes
import pandas as pd                                 # Usada para trabalhar com dataframes/tabelas
import xlrd
import yaml                                         # Usada para ler arquivos .yml (onde estão as credências do banco de dados)
import pyodbc                                       # Usada para conectar em banco de dados SQL Server (Windows)
import pymssql                                      # Usada para conectar em banco de dados SQL Server (Linux)
import datetime as dt                               # Usada para trabalhar com datas e horas
import requests                                     # Usada para fazer requisições HTTP
import json                                         # Usada para trabalhar com objetos JSON
import os                                           # Usada para trabalhar com diretórios e arquivos (sistema operacional)
import re                                           # Usada para trabalhar com expressões regulares
from decimal import Decimal, ROUND_HALF_UP          # Usada para trabalhar com números decimais e arredondar corretamente
pd.options.mode.chained_assignment = None           # Usada para retirar os avisos quando for modificar colunas do dataframe
import openpyxl
import urllib                                       # Usada para salvar arquivos nos diretorios
import xlrd                                 
import win32com.client as win32                     # Usada para abrir e navegar pelo excel
win32c = win32.constants

# Importa arquivo .xls do link
dls = 'http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls'
urllib.request.urlretrieve(dls, 'vendas-combustiveis-m3-teste.xls')

# Abre o Excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(r"C:\Users\felipe.ferreira\Documents\Python\vendas-combustiveis-m3-teste.xls")
excel.Visible = False

#Sales of oil derivative fuels by UF and product
# Definindo a tabela dinamica
pvtTable = wb.Sheets("Plan1").Range("B52").PivotTable

# Criando Dataframe de derivados do petroleo
df = pd.DataFrame(columns=['year_month', 'uf', 'product', 'unit', 'volume', 'created_at'])

# Lista com todos os produtos de derivados do petroleo
lista_produtos = []
for i in pvtTable.PivotFields('PRODUTO').PivotItems():
    lista_produtos.append(str(i))

# Lista com todos os estados 
lista_ufs = []
for i in pvtTable.PivotFields('UN. DA FEDERAÇÃO').PivotItems():
    lista_ufs.append(str(i))

# Lista com todos os meses (linhas da tabela dinamica)
lista_mes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Lista com os Anos (colunas da tabela dinamica)
lista_ano = [2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020]

# Primeiro for para correr todos os produtos de derivados do petroleo
for a in lista_produtos: 
    pvtTable.PivotFields('PRODUTO').CurrentPage = a
    # Segundo for para correr todos os Estados
    for b in lista_ufs:
        pvtTable.PivotFields('UN. DA FEDERAÇÃO').CurrentPage = b
        for i in range(len(lista_mes)): # Correr as linhas
            for j in range(len(lista_ano)): # Correr as colunas
                valor = str(wb.Sheets("Plan1").Cells(54+i,3+j).Value) # Atribui o volume atual a variavel 'valor'
                if valor not in 'None': 
                    # Criando data frame temp e concatenando cada linha já na estrutura da tabela
                    df2 = pd.DataFrame(np.array([str(lista_ano[j])+'-'+str(i+1) ,b , a.replace(' (m3)',''),a[-4:], str(valor), dt.datetime.now()]).reshape(1,6))
                    df2.columns = ['year_month', 'uf', 'product', 'unit', 'volume', 'created_at'] 
                    df = pd.concat([df, df2], axis=0)

#Sales of diesel by UF and type
# Definindo a tabela dinamica
pvtTable2 = wb.Sheets("Plan1").Range("B132").PivotTable

# Criando Dataframe de vendas de diesel
df_diesel = pd.DataFrame(columns=['year_month', 'uf', 'product', 'unit', 'volume', 'created_at'])

# Lista com todos os produtos de de vendas de diesel
lista_produtos_diesel = []
for i in pvtTable2.PivotFields('PRODUTO').PivotItems():
    lista_produtos_diesel.append(str(i))

# Lista com todos os estados 
lista_ufs_diesel = []
for i in pvtTable2.PivotFields('UN. DA FEDERAÇÃO').PivotItems():
    lista_ufs_diesel.append(str(i))

# Lista com todos os meses (linhas da tabela dinamica)
lista_mes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Lista com os Anos (colunas da tabela dinamica)
lista_ano = [2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020]

# Primeiro for para correr todos os produtos
for a in lista_produtos_diesel: 
    pvtTable2.PivotFields('PRODUTO').CurrentPage = a
    # Segundo for para correr todos os Estados
    for b in lista_ufs_diesel:
        pvtTable2.PivotFields('UN. DA FEDERAÇÃO').CurrentPage = b
    
        for i in range(len(lista_mes)): # Correr as linhas
            for j in range(len(lista_ano)): # Correr as colunas
                valor = str(wb.Sheets("Plan1").Cells(134+i,3+j).Value) # Atribui o volume atual a variavel 'valor'
                if valor not in 'None':
                    # Cria data frame temp e concatenando cada linha já na estrutura da tabela
                    df3 = pd.DataFrame(np.array([str(lista_ano[j])+'-'+str(i+1) ,b , a.replace(' (m3)',''),a[-4:], str(valor), dt.datetime.now()]).reshape(1,6))
                    df3.columns = ['year_month', 'uf', 'product', 'unit', 'volume', 'created_at']
                    df_diesel = pd.concat([df_diesel, df3], axis=0)

# Convertendo para datetime para inserir no SQL
df_diesel['year_month'] = pd.to_datetime(df_diesel['year_month'])
df['year_month'] = pd.to_datetime(df['year_month'])

try:
    # Pegando as credenciais do banco no diretorio com acesso restrito, assim as credenciais não ficam no codigo
    # with open(r'\\10.23.23.21\software_local\ProjectMaster\credentials\db_credentials.yml') as file:
    #     credencials = yaml.load(file, Loader=yaml.FullLoader)
    # file.close()

    # DB CREDENTIALS
    # server = credencials['DBINFPRD']['server']
    # database = credencials['DBINFPRD']['database']
    # username = credencials['DBINFPRD']['user'] 
    # password = credencials['DBINFPRD']['password']

    # Deixei as credenciais no codigo para que pudessem testar a resolucao do teste. (normalmente utilizaria a forma comentada acima)
    server = 'sqnprd.database.windows.net'
    database = 'DBINFPRD'
    username = 'usr_raizen'
    password = 'h4xsU7v@dg!'

    # Conectando com banco e criando cursor
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = conn.cursor()
    cursor.fast_executemany = True

    # Cria string dos inserts
    insert_to_tb_diesel = f"INSERT INTO DBINFPRD.dbo.TB_DIESEL_SALES VALUES (?,?,?,?,?,?)"
    insert_to_tb_oil = f"INSERT INTO DBINFPRD.dbo.TB_OIL_DERIVATIVES_SALES VALUES (?,?,?,?,?,?)"

    # Limpa tabela antes de inserir os dados
    cursor.execute("DELETE FROM DBINFPRD.dbo.TB_DIESEL_SALES")
    # Insere df na tabela TB_DIESEL_SALES
    cursor.executemany(insert_to_tb_diesel, df_diesel.values.tolist())     

    # Limpa tabela antes de inserir os dados
    cursor.execute("DELETE FROM DBINFPRD.dbo.TB_OIL_DERIVATIVES_SALES")
    # Insere df na tabela TB_OIL_DERIVATIVES_SALES
    cursor.executemany(insert_to_tb_oil, df.values.tolist())   
    
    # Fecha a conexão em caso de erro
    cursor.commit()
    cursor.close()
    conn.close()

except Exception as Argument:
    # Incluir tratativa de erro no imput dos dados no SQL
    print(str(Argument))
    # Fecha a conexão em caso de erro
    cursor.commit()
    cursor.close()
    conn.close()