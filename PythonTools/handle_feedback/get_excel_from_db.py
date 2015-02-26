import pymysql
import shutil,os
import win32com.client as win32
import datetime
# import pytz

# dt.replace(tzinfo=pytz.utc).astimezone(pytz.timezone('Asia/Shanghai'))

target_day = '20150208'
next_day = '20150212'

def access_mysql():
    conn = pymysql.connect(host='192.168.xxx.xxx', user='xxxuser',passwd='xxxxxx',db='xxxdb',charset="utf8")
    cur = conn.cursor()
    cur.execute("select * from user_fb where created between '20150208' and '20150209' or created between '20150210' and '20150211'")
    data=cur.fetchall()
    cur.close() #关闭游标
    conn.close() #释放数据库资源
    return data

data = access_mysql()
print("数据库数据行数：",len(data))

filename = "用户反馈" + target_day + ".xlsx"

shutil.copyfile("//Share/template/用户反馈.xlsx",os.getcwd()+"/"+filename)

xlApp = win32.gencache.EnsureDispatch('excel.application')
wb = xlApp.Workbooks.Open(os.getcwd()+"/"+filename)
ws = wb.Worksheets("所有反馈")

for i in range(0, len(data)):
    datarow = list(data[i])
    del datarow[8]
    del datarow[1]
    del datarow[0]
    for j in range(0,len(datarow)):
        # if isinstance(datarow[j],datetime.datetime):
        #     print("============datetime===========",datarow[j])
        if datarow[j]  == None:
            ws.Cells(i+2,j+1).Value = None
        else:
            ws.Cells(i+2,j+1).Value = str( datarow[j] )


wb.Save()
wb.Close()


os.startfile(os.getcwd()+"/"+filename)


# if __name__ == '__main__':
#     access_mysql()
