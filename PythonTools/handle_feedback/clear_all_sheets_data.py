import win32com.client as win32
import os

xlApp = win32.gencache.EnsureDispatch('excel.application')
wb = xlApp.Workbooks.Open(os.getcwd()+'''\用户反馈本周详情0120-0126.xlsx''')

try:
    for i in range(1,6):
        sheet = wb.Worksheets(i)
        for r in range(2,1000):
            for c in range(1,26):
                if sheet.Cells(r,c).Value != None:
                    sheet.Cells(r,c).Value = None
finally:
    wb.Save()
    wb.Close()

print('====== done ========')