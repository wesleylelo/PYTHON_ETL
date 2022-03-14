
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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="CadastroProspecção")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO cadastroprospeccao (TITULO,DATA_INI,DATA_TER, PROBABILIDADE,CORDENADOR_AREA_LIDER,AREA_LIDER) VALUES ( %s, %s, %s,%s, %s, %s)"
i = 0
while i < len(records):
    mycursor.execute(sql,list(records[i]))
    i += 1
mydb.commit()
print(mycursor.rowcount, "record inserted.")