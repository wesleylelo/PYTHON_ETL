import pandas as pd
import numpy as np

import mysql.connector 
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="Wsbolelo1998@",
  database="pyton_db_test"
)
dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="Área")

records = dados_excel.values.tolist()
print(records)
mycursor = mydb.cursor()
sql = "INSERT INTO area (TITULO,AREA_ENVOLVIDAS) VALUES ( %s, %s)"
i = 0
while i < len(records):
    if records[i][1] is np.NaN:
        records[i][1] = ""
    mycursor.execute(sql,list(records[i]))
    i += 1
mydb.commit()

print(mycursor.rowcount, "record inserted.")