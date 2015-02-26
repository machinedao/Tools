import win32com.client as win32
import os

xlApp = win32.gencache.EnsureDispatch('excel.application')
#=========================================================================================
cell_no = 17
file_path = os.getcwd()+'''\用户反馈本周详情0203-0209.xlsx'''
#=========================================================================================


wb = xlApp.Workbooks.Open(file_path,ReadOnly=1)
ws_Op = wb.Worksheets("操作问题")
ws_S = wb.Worksheets("建议反馈")
ws_Ba = wb.Worksheets("Bug_Android")
ws_BiPad = wb.Worksheets("Bug_iPad")
ws_BiPhone = wb.Worksheets("Bug_iPhone")

wbT = xlApp.Workbooks.Open(os.getcwd()+'''\用户反馈数据统计分类.xlsx''')



AllSheets = (ws_Op, ws_S, ws_Ba, ws_BiPad,ws_BiPhone)

for i in AllSheets:
    print(i.Range('A65536').End(win32.constants.xlUp).Row)

def countKeywords (a, dictKeyword):
    for r in range(2,a.Range('A65536').End(win32.constants.xlUp).Row +1 ):
        _a = a.Cells(r,12).Value
        if _a != None:
            if (_a.find(":") >= 0):
                _a = _a.split(":")[0]
            if (_a.find("：") >= 0):
                _a = _a.split("：")[0]
        if(_a in dictKeyword):
            dictKeyword[_a] += 1
        else:
            dictKeyword[_a] = 1

dOp = {}
dS = {}
dBa = {}
dBiPad = {}
dBiPhone = {}


AllKeywordDicts = [dOp,dS,dBa,dBiPad,dBiPhone]

#Bug_A


def count_sub_collections(sheetname,dicts):
    wsT = wbT.Worksheets(sheetname)
    row_count_wsT = wsT.Range('A65536').End(win32.constants.xlUp).Row
    row_count_wsT2 = wsT.UsedRange.Rows.Count
    if row_count_wsT < row_count_wsT2:
        row_count_wsT = row_count_wsT2
    for r in range(2, row_count_wsT + 1):
        if(wsT.Cells(r,1).Value in dicts):
            if(wsT.Cells(r,cell_no).Value == None):
                wsT.Cells(r,cell_no).Value = dicts[wsT.Cells(r,1).Value]
                print("wrote:",wsT.Cells(r,1).Value, wsT.Cells(r,cell_no).Value)
            del dicts[wsT.Cells(r,1).Value]
    print("=====================================================================  %s " % sheetname)
    for k,v in dicts.items():
        print(k,v)
    print("\r\n\n")

countKeywords(ws_Ba,dBa)
countKeywords(ws_BiPad,dBiPad)
countKeywords(ws_BiPhone,dBiPhone)
countKeywords(ws_Op,dOp)
countKeywords(ws_S,dS)

count_sub_collections("用户操作",dOp)
count_sub_collections("用户建议",dS)
count_sub_collections("Bug_A",dBa)
count_sub_collections("Bug_iPad",dBiPad)
count_sub_collections("Bug_iPhone",dBiPhone)


# for i in AllKeywordDicts:
#     for m,n in i.items():
#         print(m,n)

wb.Close()
wbT.Save()
wbT.Close()

os.startfile(os.getcwd()+'''\用户反馈数据统计分类.xlsx''')
print("done")
