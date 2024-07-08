import mysql.connector
mydb = mysql.connector.connect(
    host = "localhost",
    user = "root",
    passwd = "Adnan1996",
    database = "instrument_store"
)
print(mydb)
mycursor = mydb.cursor()

# mycursor.execute("CREATE DATABASE STORE")
# mycursor.execute("SHOW DATABASES")

# mycursor.execute("SHOW TABLES")
# for tb in mycursor:
#     print(tb)
mycursor.execute("SELECT * FROM instrument_store.stocks")

for items in mycursor:
    print(items)

 