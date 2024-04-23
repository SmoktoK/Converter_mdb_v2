# Run ONLY Python 32-bit!!!
import csv
import os
import pandas as pd
import pyodbc as sqlMS
import warnings

warnings.filterwarnings('ignore')


x = []

try:
    os.remove("table_list.csv")
except:
    pass

db = input('Enter name *.mdb file: ')
dbx = db
db = db + '.mdb'
dirname = os.getcwd()
if not os.path.isdir('out_csv'):
    os.mkdir('out_csv')

dir_out = os.path.join(dirname, 'out_csv')

if not os.path.isdir(os.path.join(dir_out, dbx)):
    os.mkdir(os.path.join(dir_out, dbx))
dir_out = os.path.join(dir_out, dbx)
file_path = os.path.join(dirname, db)
# print(file_path)
# print(dir_out)

driver = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
path = 'DBQ=' + file_path

connStr = driver + path

conn = sqlMS.connect(connStr)  # создать соединение с БД
cursor = conn.cursor()
for i in cursor.tables(tableType='TABLE'):
    with open('table_list.csv', 'a') as f:
        write = csv.writer(f)
        write.writerow([i.table_name])



with open('table_list.csv', 'r') as f:
    for i in f:
        i = i[:-1]
        if len(i) > 1:
            x.append(i)

for i in x:
    df = pd.read_sql(f'SELECT * FROM {i}', conn)

    df.to_csv(os.path.join(dir_out, f'{i}.csv'), index=False)





cursor.close()
del cursor
conn.close()

