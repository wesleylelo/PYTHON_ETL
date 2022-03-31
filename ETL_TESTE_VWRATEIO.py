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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="vwRateioMensalColaborador")
dados_excel.fillna(NULL, inplace = True)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO dim_vwrateiomensal (NEGOCIO,FILIAL,MERCADO,AREA_NEGOCIO,UNIDADE_NEGOCIO,PROJETO,SUB_MODALIDADE,CODIGO_CENTRO_RESULTADO,DESCRICAO_CENTRO_RESULTADO,CENTRO_RESULTADO,MATRICULA,COLABORADOR,TIPO_COLABORADOR,TIPO_CONTRATO,ANO,MES,RATEIO_PERCENTUAL) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s, %s,%s,%s,%s,%s)"
sql1 = "SELECT * FROM dim_vwrateiomensal WHERE NEGOCIO = %s AND FILIAL = %s AND MERCADO = %s AND AREA_NEGOCIO = %s AND UNIDADE_NEGOCIO = %s AND PROJETO = %s AND SUB_MODALIDADE = %s AND CODIGO_CENTRO_RESULTADO = %s AND DESCRICAO_CENTRO_RESULTADO = %s AND CENTRO_RESULTADO = %s AND MATRICULA = %s AND COLABORADOR = %s AND TIPO_COLABORADOR = %s AND TIPO_CONTRATO = %s AND ANO = %s AND MES = %s AND RATEIO_PERCENTUAL = %s"

i = 0
while i < len(records):
    print(records[i])
    records[i][11].upper()
    mycursor.execute(sql1, records[i])
    fetch = mycursor.fetchall()
    print(type(fetch))

    if(len(fetch) == 0):
      mycursor.execute(sql,records[i])
    i +=1
mydb.commit()
print(i)
print(mycursor.rowcount, "record inserted.")