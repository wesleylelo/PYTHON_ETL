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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="vwRateioMensalColaborador")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO dim_vwrateiomensal (NEGOCIO,FILIAL,MERCADO,AREA_NEGOCIO,UNIDADE_NEGOCIO,PROJETO,SUB_MODALIDADE,CODIGO_CENTRO_RESULTADO,DESCRICAO_CENTRO_RESULTADO,CENTRO_RESULTADO,MATRICULA,COLABORADOR,TIPO_COLABORADOR,TIPO_CONTRATO,ANO,MES,RATEIO_PERCENTUAL) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s, %s,%s,%s,%s,%s)"
truncate = "truncate table dim_VWRATEIOMENSAL"
mycursor.execute(truncate)
i = 0
while i < len(records):
    print(records[i])
    if records[i][5] is np.NaN:
        records[i][5] = ""
    records[i][11].upper()
    mycursor.execute(sql,records[i])
    i +=1
mydb.commit()
print(mycursor.rowcount, "record inserted.")