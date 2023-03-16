#Carga desde SQL

import pandas as pd
import pyodbc as pyodbc
from openpyxl import load_workbook


cnxn = pyodbc.connect(driver='{SQL Server Native Client 11.0}',
                      host='(local)', database='Parcial'
                      ,trusted_connection='yes')

sql = "SELECT * FROM monitor_report WHERE [Monitor_Date] = DATEADD(day, -1, convert(date, GETDATE()))"
data = pd.read_sql(sql, cnxn)

df_ayer = pd.DataFrame(data)
print(df_ayer)

df_actualizado = pd.read_excel(r'C:\Users\GW\Desktop\Python ETL Certus\Data Ops\Monitor_report_diario.xlsx')

with pd.ExcelWriter(r'C:\Users\GW\Desktop\Python ETL Certus\Data Ops\Monitor_report_diario.xlsx', mode='w') as writer:
    df_ayer.to_excel(writer, index=False)


#print(df_actualizado)



df_ayer.to_excel(r'C:\Users\GW\Desktop\Python ETL Certus\Data Ops\Monitor_report_diario.xlsx',mode='a',index=False,header=False)
df_actualizado = pd.read_excel(r'C:\Users\GW\Desktop\Python ETL Certus\Data Ops\Monitor_report_diario.csv')
print(df_actualizado)


