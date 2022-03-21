from asyncio.windows_events import NULL
from datetime import date, datetime
import pandas as pd
import numpy as np

import mysql.connector 
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="Wsbolelo1998@",
  database="pyton_db_test"
)
dados_excel = pd.read_excel("C:\\Users\\lelo0\\Downloads\\Alocação21_ELETRÔNICA.xlsm", sheet_name="BaseDados")
dados_excel_2 = dados_excel.iloc[:,np.r_[7:21]]
dados_excel_3 = dados_excel.iloc[:,np.r_[1:7]]
alocacoes = dados_excel_2.values.tolist()
alocacoes_2 = dados_excel_3.values.tolist()
mycursor = mydb.cursor()
i = len(alocacoes) - 1 
print(i)
j = 13
ii = 0
jj = 0
control = 0
sql = "INSERT INTO fato_base_dados (ATIVIDADES, TITULO,CATEGORIAS,MEMBROS, PROBABILIDADE, CARGO , DATA_, ALOCACAO) VALUES (%s, %s, %s,%s, %s, %s,%s, %s)"
while ii < i:
    jj = 0
    control = 0
    if (alocacoes_2[ii][5]) is np.nan:
            alocacoes_2[ii][5] = ""
    if np.isnan(alocacoes_2[ii][4]):
            alocacoes_2[ii][4] = NULL
    else:
        alocacoes_2[ii][4] = alocacoes_2[ii][4]
    while jj < j :  
        if (alocacoes[ii][jj]) > 0:
            aloc = alocacoes_2[ii]
            if(jj == 0):
                if control == 0:
                    aloc.append(datetime(2021, 11,1).date()) 
                    aloc.append(alocacoes[ii][jj])
                    control = 1
                else:
                    aloc.pop()
                    aloc.pop()
                    aloc.append(datetime(2021, 11,1).date()) 
                    aloc.append(alocacoes[ii][jj])
            elif jj == 1:
                if control == 0:
                    aloc.append(datetime(2021, 12, 1).date()) 
                    aloc.append(alocacoes[ii][jj])
                    control = 1
                else:
                    aloc.pop()
                    aloc.pop()
                    aloc.append(datetime(2021, 12,1).date()) 
                    aloc.append(alocacoes[ii][jj])
            else:
                if control == 0:
                    aloc.append(datetime(2022, jj - 1, 1).date())
                    aloc.append(alocacoes[ii][jj])
                    control = 1
                else:
                    aloc.pop()
                    aloc.pop()
                    aloc.append(datetime(2022, jj - 1,1).date()) 
                    aloc.append(alocacoes[ii][jj])
            print(aloc)
            mycursor.execute(sql,aloc)
        jj += 1
    ii += 1
mydb.commit()
print(mycursor.rowcount, "record inserted.")
            
