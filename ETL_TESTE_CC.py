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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="CadastroCC")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO cadastrocc (TITULO,DATA_INI,DATA_TER, PROBABILIDADE, LIDER_TECNICO , AREA_LIDER, COORDENADOR_AREA_LIDER, CUSTO_A,CUSTO_B_C,CUSTO_C,ORGAO_FOMENTO,EMPRESA) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s)"
i = 0
while i < len(records):
    if (records[i][7] is pd.NA or np.NAN):
        records[i][7] = NULL
    if (records[i][8] is pd.NA or np.NAN):
        records[i][8] = NULL
    if (records[i][9] is pd.NA or np.NAN):
        records[i][9] = NULL
    mycursor.execute(sql,list(records[i]))
    i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")