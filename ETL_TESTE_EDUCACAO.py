
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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="CadastroEducação")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO cadastroeducacao (TITULO,DATA_INI,DATA_TER, PROBABILIDADE, COORDENADOR, CATEGORIA, DIA, HORARIO, CH) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s)"
i = 0
while i < len(records):
    if (records[i][1] is pd.NaT):
        records[i][1] = NULL
    if (records[i][2] is pd.NaT):
        records[i][2] = NULL
    if (records[i][4] is pd.NA or np.NAN):
        records[i][4] = ""
    if (records[i][5] is pd.NA or np.NAN):
        records[i][5] = ""
    if (records[i][6] is pd.NA or np.NAN):
        records[i][6] = NULL
    if (records[i][7] is pd.NA or np.NAN):
        records[i][7] = ""
    if (records[i][8] is pd.NA or np.NAN):
        records[i][8] = ""
    mycursor.execute(sql,list(records[i]))
    i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")