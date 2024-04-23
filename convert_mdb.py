import os
import pyodbc
import csv
import pandas_access as mdb
import pandas as pd
from sys import platform
import warnings
import pyodbc as sqlMS

warnings.filterwarnings('ignore')

x = []


def create_cfg():  # Generate config and driver for work with *.mdb file
    db = input('Enter name *.mdb file: ')
    db = db + '.mdb'
    x = 'find / -name libmdbodbc.so'
    if 'installed' in (os.popen('apt search odbc-mdbtools').read()):
        print('Driver already installed!!!')
    else:
        os.system('sudo apt install odbc-mdbtools')
        os.system('rm /etc/odbcinst.ini')
        print('Driver install complite!!')

        out = os.popen(x).read()

        with open('odbcinst.ini', 'r') as f:
            lines = f.readlines()
            for i in lines:
                if 'Driver=/usr/' in i:
                    i = f'Driver={out}'
                with open('/etc/odbcinst.ini', 'a') as f_new:
                    f_new.write(i)
        print('Config generated!')
    create_csv(db)


def create_csv(db):  # Generate csv from *.mdb file

    folder = db.replace('.mdb', '')
    dirname = os.getcwd()
    dirname = os.path.join(dirname, folder)
    print(f'Folder {folder} created!')

    if not os.path.isdir(folder):
        os.mkdir(folder)

    try:
        os.system('rm table_list.csv')
        os.system('rm path.csv')


    except:
        pass

    x = mdb.list_tables(db)
    x[-1] = x[-1].replace('\n', '')

    with open('table_list.csv', 'w') as f:
        write = csv.writer(f)
        write.writerow(x)
        print('Table list created!')

    conn = pyodbc.connect(r'Driver={MDBToolsODBC};'f"DBQ={db};")
    cursor = conn.cursor()

    with open('table_list.csv', 'r') as f:
        for i in f:
            for x in i.split(','):
                x = x.replace('\n', '')
                filepath = os.path.join(dirname, f'{x}.csv')
                with open('path.csv', 'a') as data:
                    writer = csv.writer(data)
                    writer.writerow([filepath])

                cursor.execute(f'SELECT * FROM {x}')
                row = cursor.fetchall()
                if len(row) == 0:
                    row = ' '

                for z in row:
                    with open(filepath, 'a') as f:
                        # print(filepath)
                        writer = csv.writer(f)
                        writer.writerow(z)

        print('Add data complite!')

    csv_to_df()


def csv_to_df():
    dictionary = {}
    line_x = ['FIND_KEY', 'DATA_VAL']
    dirname = os.getcwd()
    dirname = os.path.join(dirname, 'EM133')

    with open('table_list.csv', 'r') as f:
        for i in f:
            for x in i.split(','):
                x = x.replace('\n', '')
                filepath = os.path.join(dirname, f'{x}.csv')
                df = pd.read_csv(filepath, on_bad_lines='skip', names=line_x)
                dictionary[x] = df

    print(dictionary)



# Run ONLY Python 32-bit!!!
def win_converter():
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


if platform == "linux" or platform == "linux2":
    print('Nix system')
    create_cfg()
elif platform == "darwin":
    print('OS X')
elif platform == "win32":
    print('Win system')
    win_converter()
