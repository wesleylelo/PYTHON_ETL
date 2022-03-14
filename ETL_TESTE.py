import pandas as pd
import mysql.connector 
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="Wsbolelo1998@",
  database="pyton_db_test"
)
dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="GP")
records = dados_excel.to_records(index=False)
mycursor = mydb.cursor()
sql = "INSERT INTO gp (TITULO,GERENTE_PROJETO) VALUES ( %s, %s)"
i = 0
while i < len(records):
    mycursor.execute(sql,tuple(records[i]))
    i += 1
mydb.commit()
print(records)
print(mycursor.rowcount, "record inserted.")



