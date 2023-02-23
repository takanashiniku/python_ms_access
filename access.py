import pyodbc as pyo
import pandas as pd
import openpyxl

#新規エクセル作成
url = 'C:\\Users\\tomon\\OneDrive\\デスクトップ\\実験中\\ch.xlsx'

#アクセス接続
conncect = (
	'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
	'DBQ=C:\\Users\\tomon\\OneDrive\\デスクトップ\\賃金台帳(L05) .mdb;'
	)

# Connect Access DB Stats.accdb
conncect = pyo.connect(conncect)

sql = 'SELECT * FROM Chingin1_Pato'

dataframe = pd.read_sql(sql, conncect)
conncect.close()

ep = dataframe.loc[dataframe["給与計算年月"] == "202210"]

with pd.ExcelWriter(url) as writer:
    ep.to_excel(writer, sheet_name='eria1',index=False)
    ep.to_excel(writer, sheet_name="eria2",index=False)

