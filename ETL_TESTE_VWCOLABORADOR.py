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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="vwColaborador")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()

sql = "INSERT INTO dim_vwcolaborador (MATRICULA,CPF,NOME, FILIAL_LOTACAO, MERCADO_LOTACAO , AREA_TECNOLOGICA_LOTACAO, UNIDADE_NEGOCIO_LOTACAO, FORMACAO,TIPO_COLABORADOR,CLASSE_COLABORADOR,EMAIL, ID_CARGO ,CARGO, DES_STATUS, CATEGORIA_COLABORADOR, DATA_ADMISSAO, DATA_DEMISSAO, TIPO_CONTRATO) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s, %s,%s,%s, %s,%s,%s)"
i = 0
while i < len(records):
  print(records[i])
  if records[i][1] is np.NaN:
    records[i][1] = str('')
  else:
    records[i][1] = str(records[i][1])[:-2]
  if records[i][16] is np.nan:
    records[i][16] = NULL
  else:
    records[i][16] = records[i][16]
  if np.isnan(records[i][11]):
    records[i][11] = ""
  else:
    records[i][11] = records[i][11]
  if (records[i][8]) is np.NaN:
    records[i][8] = ""
  else:
    records[i][8] = records[i][8]
  if (records[i][9]) is np.NaN:
    records[i][9] = ""
  else:
    records[i][9] = records[i][9]
  if records[i][10] is np.NaN:
    records[i][10] = ""
  records[i][7] = ""
  records[i].append("")
  print(records[i])
  mycursor.execute(sql,records[i])
  i +=1
mydb.commit()
print(mycursor.rowcount, "record inserted.")
print(type(records[1][0]))
print(str(records[3][1])[:-2])