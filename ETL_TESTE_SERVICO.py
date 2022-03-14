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

dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="CadastroServiço")
print(dados_excel)
records = dados_excel.values.tolist()
mycursor = mydb.cursor()
sql = "INSERT INTO cadastroservico (TITULO,DATA_INI,DATA_TER, PROBABILIDADE, LIDER_TECNICO , AREA_LIDER, COORDENADOR_AREA_LIDER, CUSTO_A,CUSTO_B_C,CUSTO_C,ORGAO_FOMENTO,EMPRESA) VALUES ( %s, %s, %s,%s, %s, %s,%s,%s, %s,%s,%s, %s)"
i = 0
while i < len(records):
    if (records[i][1] is pd.NA or np.NAN or pd.NaT):
        records[i][1] = NULL
    if (records[i][2] is pd.NA or np.NAN or pd.NaT):
        records[i][2] = NULL
    if (records[i][5] is pd.NA or np.NAN or pd.NaT):
        records[i][5] = ""
    if (records[i][6] is pd.NA or np.NAN or pd.NaT):
        records[i][6] = ""
    if (records[i][7] is pd.NA or np.NAN or pd.NaT):
        records[i][7] = NULL
    if (records[i][8] is pd.NA or np.NAN or pd.NaT):
        records[i][8] = NULL
    if (records[i][9] is pd.NA or np.NAN or pd.NaT):
        records[i][9] = NULL
    if (records[i][10] is pd.NA or np.NAN or pd.NaT):
        records[i][10] = ""
    if (records[i][11] is pd.NA or np.NAN or pd.NaT):
        records[i][11] = ""
    mycursor.execute(sql,list(records[i]))
    i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")