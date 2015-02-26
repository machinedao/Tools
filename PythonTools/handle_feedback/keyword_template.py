import win32com.client as win32
import shutil, os

def get_keyword_from_template():
    xlApp = win32.gencache.EnsureDispatch('excel.application')
    xlApp.Visible = False
    shutil.copyfile("//Share/用户反馈数据统计分类.xlsx",os.getcwd()+"/用户反馈数据统计分类.xlsx")
    try:
        workBookTemplate = xlApp.Workbooks.Open(os.getcwd()+"/用户反馈数据统计分类.xlsx",ReadOnly=1)
        wsOperation = workBookTemplate.Worksheets("用户操作")
        wsSuggestion = workBookTemplate.Worksheets("用户建议")
        wsBugAndroid = workBookTemplate.Worksheets("Bug_A")
        wsBugIPad = workBookTemplate.Worksheets("Bug_iPad")
        wsBugIPhone = workBookTemplate.Worksheets("Bug_iPhone")
        rowCountOperation = wsOperation.Range('A65536').End(win32.constants.xlUp).Row
        rowCountSuggestion = wsSuggestion.Range('A65536').End(win32.constants.xlUp).Row
        rowCountBugAndroid = wsBugAndroid.Range('A65536').End(win32.constants.xlUp).Row
        rowCountBugIPad = wsBugIPad.Range('A65536').End(win32.constants.xlUp).Row
        rowCountBugIPhone = wsBugIPhone.Range('A65536').End(win32.constants.xlUp).Row

        print("%d %d %d %d %d" %(rowCountOperation, rowCountSuggestion, rowCountBugAndroid,rowCountBugIPad,rowCountBugIPhone))

        dictOperation = {}
        dictSuggestion = {}
        dictBugAndroid = {}
        dictBugIPad = {}
        dictBugIPhone = {}

        def keyword_from_excel(rowCount,ws,d):
            for i in range(1, rowCount):
                d[ws.Cells(i,1).Value] = 0

        keyword_from_excel(rowCountOperation,wsOperation,dictOperation)
        keyword_from_excel(rowCountSuggestion,wsSuggestion,dictSuggestion)
        keyword_from_excel(rowCountBugAndroid,wsBugAndroid,dictBugAndroid)
        keyword_from_excel(rowCountBugIPad,wsBugIPad,dictBugIPad)
        keyword_from_excel(rowCountBugIPhone,wsBugIPhone,dictBugIPhone)

        # for k,v in dictSuggestion.items():
        #     print(k,v)
        dicts = (dictOperation,dictSuggestion,dictBugAndroid,dictBugIPad,dictBugIPhone)

        if not (rowCountOperation-1 == len(dicts[0]) and rowCountSuggestion-1 == len(dicts[1]) and rowCountBugAndroid-1 == len(dicts[2])
            and rowCountBugIPad-1 == len(dicts[3]) and rowCountBugIPhone-1 == len(dicts[4])):
            raise EOFError("读取文档时有错误发生，读取不完整，请检查")
    finally:
        workBookTemplate.Close()

    print(len(dicts[1]))
    return dicts

