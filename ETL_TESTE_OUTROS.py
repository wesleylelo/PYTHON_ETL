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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="CadastroOutros")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO cadastrooutros (TITULO,DATA_INI,DATA_TER, PROBABILIDADE, COORDENADOR, COLUNA) VALUES ( %s, %s, %s,%s, %s, %s)"
i = 0
while i < len(records):
    if records[i][4] is pd.NA or np.NAN:
        records[i][4] = ""
    records[i].append("")
    mycursor.execute(sql,list(records[i]))
    i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")