import win32com.client as win32
import os,shutil

week_tag = "0203-0209"


xlApp = win32.gencache.EnsureDispatch('excel.application')

op_weekly = []
su_weekly = []
bug_weekly = []
bug_a_weekly = []
bug_iphone_weekly = []
bug_ipad_weekly = []

def get_excel_data(wb,sheetname, datalist):
    _ws = wb.Worksheets(sheetname)
    _rowcount = _ws.UsedRange.Rows.Count    #Range('A65536').End(win32.constants.xlUp).Row
    _rowcount2 = _ws.Range('A65536').End(win32.constants.xlUp).Row
    if _rowcount < _rowcount2:
        _rowcount = _rowcount2
    print("=======================================================================================",sheetname,_rowcount)
    for r in range(2, _rowcount+1):
        _templist = []
        for i in range(1,14):
            _templist.append(_ws.Cells(r,i).Value)
        datalist.append(_templist)


filenames = []

for filename in os.listdir(os.getcwd()+"/resources"):
    filenames.append(filename)


for filename in filenames:
    filepath = os.getcwd() + '/resources/' + filename
    wb = xlApp.Workbooks.Open(filepath, ReadOnly=1)
    get_excel_data(wb,"操作问题",op_weekly)
    get_excel_data(wb,"建议反馈",su_weekly)
    get_excel_data(wb,"Bug反馈",bug_weekly)
    wb.Close()

for i in (op_weekly,su_weekly,bug_weekly):
    print(len(i))


for i in range(0,len(bug_weekly)):
    if bug_weekly[i][6].find("ipadiap")>=0:
        bug_ipad_weekly.append(bug_weekly[i])
    elif bug_weekly[i][6].find("iphoneiap")>=0:
        bug_iphone_weekly.append(bug_weekly[i])
    elif bug_weekly[i][6].find("android")>=0:
        bug_a_weekly.append(bug_weekly[i])
    else:print("==========pls check, not Android, not iPad, not iPhone"+"=========="+bug_weekly[i][6])

print(len(bug_a_weekly),len(bug_ipad_weekly),len(bug_iphone_weekly))

filename = "用户反馈本周详情" + week_tag + ".xlsx"
shutil.copyfile("//Share/template/用户反馈本周详情.xlsx",os.getcwd()+"/"+filename)
weeklydetals = xlApp.Workbooks.Open(os.getcwd()+'/'+filename)

bugaws = weeklydetals.Worksheets("Bug_Android")
bugiphonews = weeklydetals.Worksheets("Bug_iPhone")
bugipadws = weeklydetals.Worksheets("Bug_iPad")

def write_data_to_weeklydetail(wb,sheetname,datalist):
    _ws = wb.Worksheets(sheetname)
    for r in range(2,len(datalist)+1):
        for c in range(1,14):
            if _ws.Cells(r,c).Value == None:
                _ws.Cells(r,c).Value = datalist[r-2][c-1]
            else:
                print("==================target excel cell is not None, pls check===================",sheetname,r,c)


write_data_to_weeklydetail(weeklydetals,"操作问题",op_weekly)
write_data_to_weeklydetail(weeklydetals,"建议反馈",su_weekly)
write_data_to_weeklydetail(weeklydetals,"Bug_Android",bug_a_weekly)
write_data_to_weeklydetail(weeklydetals,"Bug_iPad",bug_ipad_weekly)
write_data_to_weeklydetail(weeklydetals,"Bug_iPhone",bug_iphone_weekly)

weeklydetals.Save()
weeklydetals.Close()

print("done")