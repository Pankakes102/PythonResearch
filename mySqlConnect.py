from sqlalchemy import *

#config = {user='root',password='vmdb',host='localhost:3306',database='gis'}
engine = create_engine('mysql+pymysql://root:vmdb@localhost:3306/gis')
connection = engine.connect()
query = connection.execute('select * from info')
for q in query:
	print('name:',q['uniqueId'])
connection.close()


# connection = mysql.connection.connect(user='root',password='vmdb',host='localhost:3306',database='gis')
# cursor = connection.cursor()
# query = 'select * from info'
# cursor.execute(query)
# for (name) in cursor:
# 	print(name)
# connection.close()

