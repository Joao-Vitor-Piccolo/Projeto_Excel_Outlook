import sqlite3 as sql

tabela = sql.connect('database.db')
cursor = tabela.cursor()


cursor.execute('SELECT * FROM dados')
registros1 = cursor.fetchall()
print(registros1)
print(registros1[0][0])