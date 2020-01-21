import pyodbc

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=D:\Temp\db\parts.mdb;')
conn.autocommit = True
cursor = conn.cursor()

param = 'Johnson Controls'
sql=("""
                select partnr, purchaseprice1, salesprice1, lastchange from tblPart
                where manufacturer=?
                """)

cursor.execute(sql,param)





# find items in price table which have in db of eplan


t1 = [['A99DY-200C','55555'],['MS-NAE3510-2', '555555']]


for x in t1:
    x1= x[1]
    x2= x[0]   
    sql_update_request ="update tblPart set salesprice1=? where partnr=?"
    conn.execute(sql_update_request,x1,x2)
    conn.commit

row = cursor.fetchall()
for i in row:
    print(i)

conn.close

