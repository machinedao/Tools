import win32com.client as win32
import os
from check_mail_phone import isPN

xlApp = win32.gencache.EnsureDispatch('excel.application')

wb = xlApp.Workbooks.Open(os.getcwd()+'''\用户反馈20150208.xlsx''')

wsallfd = wb.Worksheets("所有反馈")
wsmsg = wb.Worksheets("短信记录单")

row_count_allfb = wsallfd.Range('A65536').End(win32.constants.xlUp).Row
row_count_allfb2 = wsallfd.UsedRange.Rows.Count
if row_count_allfb < row_count_allfb2:
    row_count_allfb = row_count_allfb2

# row_count_msg = wsmsg.Range('A65536').End(win32.constants.xlUp).Row

messages = []
collectiondict = {}

for r in range(2,row_count_allfb + 1):
    if wsallfd.Cells(r,14).Value != None and wsallfd.Cells(r,4).Value != None:
        if isPN(wsallfd.Cells(r,4).Value):
            messages.append([wsallfd.Cells(r,4).Value,wsallfd.Cells(r,14).Value])
    collectionname = wsallfd.Cells(r,1).Value
    if collectionname != None:
        rowlist = []
        for c in range(1,14):
            rowlist.append(wsallfd.Cells(r,c).Value)
        if collectionname in collectiondict:
            collectiondict[collectionname].append(rowlist)
        else:
            collectiondict[collectionname] = [rowlist,]



for i in range(0,len(messages)):
    wsmsg.Cells(i+2,1).Value = messages[i][0]
    wsmsg.Cells(i+2,2).Value = messages[i][1]


# for i in messages:
#     print(i)

def cp_to_subsheet(collectionname,sheetname):
    ws = wb.Worksheets(sheetname)
    for k,v in collectiondict.items():
        if k == sheetname:
            for r in range(0,len(collectiondict[k])):
                for c in range(0,len(collectiondict[k][r])):
                    ws.Cells(r+2,c+1).Value = collectiondict[k][r][c]



cp_to_subsheet("操作问题","操作问题")
cp_to_subsheet("建议反馈","建议反馈")
cp_to_subsheet("Bug反馈","Bug反馈")


# for k,v in collectiondict.items():
#     print(k,v)


wb.Save()
wb.Close()
print("========= done =========")