from asyncio.windows_events import NULL
import pandas as pd
import numpy as np
import mysql.connector 

mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="Wsbolelo1998@",
  database="pyton_db_test"
)

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="vwColaborador")
print(dados_excel)
dados_excel.fillna(NULL, inplace = True)

mycursor = mydb.cursor()
mycursor.execute("SELECT * FROM dim_vwrateiomensal")
dados_sql = pd.DataFrame(mycursor.fetchall())
print(dados_sql)
dados_merge = dados_excel.merge(dados_sql,"left", left_on="Nome", right_on=12).iloc[:,np.r_[:17]]
print(dados_merge)
dados_merge.drop_duplicates(subset=["Nome"])
dados_merge1 = pd.DataFrame(dados_merge).query("descStatus == 'Ativo' and CategoriaColaborador  == 'Efetivo'")
print(dados_merge1)
sql = "INSERT INTO dim_vwcolaborador (MATRICULA,CPF,NOME, FILIAL_LOTACAO, MERCADO_LOTACAO , AREA_TECNOLOGICA_LOTACAO, UNIDADE_NEGOCIO_LOTACAO, FORMACAO,TIPO_COLABORADOR,CLASSE_COLABORADOR,EMAIL, ID_CARGO ,CARGO, DES_STATUS, CATEGORIA_COLABORADOR, DATA_ADMISSAO, DATA_DEMISSAO, TIPO_CONTRATO) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s, %s,%s,%s, %s,%s,%s)"
sql1 = "SELECT * FROM dim_vwcolaborador WHERE MATRICULA = %s AND CPF = %s AND NOME = %s AND FILIAL_LOTACAO = %s AND MERCADO_LOTACAO = %s AND AREA_TECNOLOGICA_LOTACAO = %s AND UNIDADE_NEGOCIO_LOTACAO = %s AND FORMACAO = %s AND TIPO_COLABORADOR = %s AND CLASSE_COLABORADOR = %s AND EMAIL = %s AND ID_CARGO = %s AND CARGO = %s AND DES_STATUS = %s AND CATEGORIA_COLABORADOR = %s AND DATA_ADMISSAO = %s AND DATA_DEMISSAO = %s AND TIPO_CONTRATO = %s"
records = dados_merge1.values.tolist()
i = 0
while i < len(records):
  records[i][1] = str(records[i][1])[:-2]
  records[i][7] = ""
  records[i][12].upper()
  records[i].append("ESTAGIÁRIO")
  print(records[i])
  mycursor.execute(sql1, records[i])
  fetch = mycursor.fetchall()
  print(type(fetch))

  if(len(fetch) == 0):
    mycursor.execute(sql,records[i])
  i +=1
mydb.commit()
print(mycursor.rowcount, "record inserted.")