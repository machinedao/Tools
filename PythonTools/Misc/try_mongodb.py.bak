import pymongo

connection=pymongo.Connection('192.xxx.xx.xx',27000)

db = connection.temp
collection = db.Collections
clicks = db.click_bak


#print(db.collection_names()) # 获取所有collection（相当于SQL的show tables）


for click in clicks.find({"deviceType":"android"}).limit(10):
    print(click)



connection.close()
