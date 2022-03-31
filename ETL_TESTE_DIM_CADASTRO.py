from asyncio.windows_events import NULL
from mysqlx import ColumnType
import pandas as pd
import numpy as np

import mysql.connector 
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="Wsbolelo1998@",
  database="pyton_db_test"
)
data_projeto = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroProjeto")
data_projeto.insert(0,column="Name", value="CadastroProjeto")
data_projeto.drop(columns=["CUSTO A","CUSTO B+C", "CUSTO C", "ÓRGÃO DE FOMENTO", "EMPRESA", "ÁREA LÍDER", "COORDENADOR ÁREA LÍDER"], inplace=True)
data_projeto.insert(6,column="Cadastro", value="Projeto")
data_projeto.rename(columns={'TÍTULO':'Atividade'},inplace=True)
data_prospeccao = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroProspecção")
data_prospeccao.insert(0,column="Name", value="CadastroProspecção")

data_prospeccao.drop(columns=["ÁREA LÍDER"], inplace=True)
data_prospeccao.insert(6,column="Cadastro", value="Prospecção")
data_prospeccao.rename(columns={'TÍTULO':'Atividade'},inplace=True)

data_outros = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroOutros")
data_outros.insert(0,column="Name", value="CadastroOutros")
data_outros.insert(6,column="Cadastro", value="Outros")
data_outros.rename(columns={'TÍTULO':'Atividade'},inplace=True)

data_gestao = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroGestão")
data_gestao.insert(0,column="Name", value="CadastroGestão")
data_gestao.drop(columns=["ÁREA LÍDER"], inplace=True)
data_gestao.insert(6,column="Cadastro", value="Gestão")
data_gestao.rename(columns={'TÍTULO':'Atividade'},inplace=True)

data_servico = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroServiço")
data_servico.insert(0,column="Name", value="CadastroServiço")
data_servico.drop(columns=["CUSTO A","CUSTO B+C", "CUSTO C", "ÓRGÃO DE FOMENTO", "EMPRESA", "ÁREA LÍDER", "COORDENADOR ÁREA LÍDER"], inplace=True)
data_servico.insert(6,column="Cadastro", value="Serviço")
data_servico.rename(columns={'TÍTULO':'Atividade'},inplace=True)

data_educacao = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroEducação")
data_educacao.insert(0,column="Name", value="CadastroEducação")
data_educacao.drop(columns=["CATEGORIA","DIA","HORÁRIO","CH"], inplace=True)
data_educacao.insert(6,column="Cadastro", value="Educação")
data_educacao.rename(columns={'TÍTULO':'Atividade'},inplace=True)

data_cc = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroCC")
data_cc.insert(0,column="Name", value="CadastroCC")
data_cc.drop(columns=["CUSTO A","CUSTO B+C", "CUSTO C", "ÓRGÃO DE FOMENTO", "EMPRESA", "ÁREA LÍDER", "COORDENADOR ÁREA LÍDER"], inplace=True)
data_cc.insert(6,column="Cadastro", value="CC")
data_cc.rename(columns={'TÍTULO':'Atividade'},inplace=True)

data_all = pd.concat([data_projeto,data_servico,data_cc, data_gestao, data_prospeccao, data_educacao, data_outros]).iloc[:,np.r_[:7]]
data_finish = data_all.fillna(NULL)
records = data_finish.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO dim_cadastro (NAME,ATIVIDADE,DATA_INI,DATA_TER,PORBABILIDADE,LIDER_ATIVIDADE,CADASTRO) VALUES (%s, %s, %s,%s, %s, %s,%s)"
sql1 = "SELECT * FROM dim_cadastro WHERE NAME = %s AND ATIVIDADE = %s AND DATA_INI = %s AND DATA_TER = %s AND PORBABILIDADE = %s AND LIDER_ATIVIDADE = %s AND CADASTRO = %s"
i = 0
while i < len(records):
  mycursor.execute(sql1, records[i])
  fetch = mycursor.fetchall()
  if(len(fetch) == 0):
    print(records[i])
    mycursor.execute(sql,records[i])
  i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")