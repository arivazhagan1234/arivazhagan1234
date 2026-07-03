import sqlite3

#create a db file
con=sqlite3.connect("mybio.db")
cursor=con.cursor()
#Drop table 
#cursor.execute("drop table if exists bio")
#create a table
cursor.execute(""" create table if not exists bio (id integer primary key, name text, age integer, gender text)""")

#store data in the table
cursor.execute("insert into bio(name, age, gender) values(?, ?, ?)", ('Ariva', 23, 'male'))  
con.commit()

cursor.execute("PRAGMA table_info(bio)")
rows=cursor.fetchall()
print("Table Info:",rows)
exists={row[1] for row in rows}
print("Column Names:", exists)
newdata={'cell_num': 'Numeric', 'Address': 'Text', 'Email': 'Text'}
for data, defination in newdata.items():
    if data not in exists:
        cursor.execute(f'alter table bio add column {data} {defination}')

cursor.execute("PRAGMA table_info(bio)")
print("Table Info:",cursor.fetchall())
cursor.execute('insert into bio (name,age,gender,cell_num,Address,Email) values(?,?,?,?,?,?)', ('Tamil', 24,'female','8237283723387','ari@gmal.com','A kumaramagalam'))
con.commit()
print("Data Inserted Successfully")
cursor.execute("select * from bio") 
print("Data in Table:",cursor.fetchall())
cursor.execute("select count(*) from bio where name='Ariva'")
print("Count:",cursor.fetchone()[0])
con.close()
