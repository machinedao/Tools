import win32com.client as win32
import os

xlApp = win32.gencache.EnsureDispatch('excel.application')
# xlApp.Visible = False

# -------------------- 更改路径和文件名 --------------------
file_path = os.getcwd() + '\用户反馈20150208.xlsx'
# ----------------------------------------------------------

'''
结束后，请注意检查excel，智能分类错误时请手工更正子类别
只有 类型，备注，处理人 三列都为空的行，才会被自动分类

filters规则：
类型A，后面只有两个元素组，excel中的设备类型匹配到第二个元素组中的任何一项，则把该行的第一格改成第一个元组中的内容
类型B，后面跟2个以上元素组(无个数限制)，除了匹配设备类型，还会匹配excel中的内容一列
其余，第一个元素组中没有标注类型，只匹配excel中的内容一列，元素组之间是and关系，组内子元素之间是or关系
例如("购","买","付款","付过款"),("没有","找不到","看不")，相当于（购or买or付款or付过款）and（没有or找不到or看不），
excel中的“买了...看不到”“...付过款了...找不到...”就会被匹配到
而方括号[]取非，例如("操作问题","笔记划线操作"),("笔记","划线"),["印象笔记","有道云笔记"])
相当于("笔记" or "划线") and not ["印象笔记" or "有道云笔记"]，
只有该行中含有“笔记、划线”任一词，并且不包含"印象笔记、有道云笔记"任一词才会被匹配。

'''

filters = (
    (("A",),("PC用户",),("pc",)), #注意括号中只有一个元素时，后面加逗号才会被正确识别为元组
    (("B",),("操作问题","购买支付"),('ip',),("信用卡","交易失败")),
    (("B",),("操作问题","礼券使用"),('android',),("券","劵","礼卷","礼品卷"),["证券",]),
    (("B",),("建议反馈","微信登录"),('ip',),("微信",)),
    (("B",),("Bug反馈","微信登录"),('ip',),("微信",),("登",),("不","失败","响应","反应"),["支付","付款"]),
    (("Bug反馈","应用崩溃"),("闪退","崩溃","死机","卡住了","秒退")),
    (("资源问题","签到失败"),("签到","补签","簽到"),("失败","不成功","没法","不能","不了","无法","出错","错误")),
    (("操作问题","第三方登录"),("第三方",),("登",)),
    (("操作问题","找回密码"),("密码",),("找回","修改","忘")),
    (("Bug反馈","登录问题"),("登不","登录失败","登陆不","登录不","登陆失败","无法登录","无法登陆","验证码错误","验证码不对","验证码不正常","登录异常","登陆异常")),
    (("Bug反馈","iphone6 plus兼容"),("iphone6",),("plus",)),
    (("建议反馈","Windowphone版本"),("window",),("phone",)),
    (("建议反馈","Windowphone版本"),("windowphone","wp系统","wp手机")),
)

workBook = xlApp.Workbooks.Open(file_path)
ws = workBook.Worksheets("所有反馈")
rowCount = ws.UsedRange.Rows.Count  #取得行数，附取得列数ActiveSheet.UsedRange.Columns.Count
rowCount2 = ws.Range('A65536').End(win32.constants.xlUp).Row
# 有时候有些excel文档上述两种取行方法某一种只能取到1行等，所以采用两种办法同时取行数
if rowCount < rowCount2:
    rowCount = rowCount2

print("【所有反馈】行数：",rowCount)

def isMatch(cellValue,keywords,flag="flag"):
    tempAnd = True
    if flag == "B" :
        startIndex = 3
    else:startIndex = 1
    for k in range(startIndex,len(keywords)):
        # print(k)
        if isinstance(keywords[k],str):
            print("Error!Please modify your filters and make sure a comma following the only one item")
        tempOr = False
        isFoundFlag = False
        for ki in range(0,len(keywords[k])):
            if isinstance(keywords[k],tuple):
                if(ws.Cells(r,2).Value.find(keywords[k][ki])>=0):
                    tempOr = True
            elif isinstance(keywords[k],list):
                # print(keywords[k],keywords[k][ki])
                if(ws.Cells(r,2).Value.find(keywords[k][ki])>=0):
                    tempOr = False
                    isFoundFlag = True
                elif (ki==len(keywords[k])-1 and isFoundFlag==False):
                    tempOr = True
            if(ki == len(keywords[k])-1 and tempOr == False):
                tempAnd = False
            elif (ki == len(keywords[k])-1 and tempOr == True):
                tempAnd = tempAnd and tempOr
    return tempAnd

try:
    for r in range(1,rowCount+1):
        for keywords in filters:
            if(keywords[0][0]=="A"):
                if(ws.Cells(r,7).Value.find(str(keywords[2][0]))>=0):
                    ws.Cells(r,1).Value = keywords[1][0]
            elif(keywords[0][0]=="B"):
                if(ws.Cells(r,7).Value.find(str(keywords[2][0]))>=0):
                    tempMatch = isMatch(ws.Cells(r,2).Value,keywords,"B")
                    # if tempMatch == True: print("-----------------------",ws.Cells(r,1).Value)
                    if tempMatch == True and (ws.Cells(r,1).Value in (None,'')):
                        ws.Cells(r,1).Value = keywords[1][0]
                        ws.Cells(r,12).Value = keywords[1][1]
                        print(r, ws.Cells(r,1).Value, ws.Cells(r,12).Value)
            else:
                isMatchK = isMatch(ws.Cells(r,2),keywords)
                # 只有 类型，备注，处理人 三列都为空的行，才会被自动分类
                if (isMatchK == True and ws.Cells(r,1).Value in (None,'') and ws.Cells(r,12).Value in (None,'')
                    and ws.Cells(r,13).Value in (None,'')):
                    ws.Cells(r,1).Value = keywords[0][0]
                    ws.Cells(r,12).Value = keywords[0][1]
                    print(r, ws.Cells(r,1).Value, ws.Cells(r,12).Value)
finally:
    workBook.Save()
    workBook.Close()
    print("============================")
print("filter collections ok\n---please check and modify your excel---\n---Next step: Send Mail---")
print("============================")

os.startfile(file_path)