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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA (1).xlsm", sheet_name="CadastroProjeto")
dados_excel.fillna(NULL, inplace = True)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO cadastroprojeto (TITULO,DATA_INI,DATA_TER, PROBABILIDADE, LIDER_TECNICO , AREA_LIDER, COORDENADOR_AREA_LIDER, CUSTO_A,CUSTO_B_C,CUSTO_C,ORGAO_FOMENTO,EMPRESA) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s)"
i = 0
while i < len(records):
    mycursor.execute(sql,records[i])
    i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")