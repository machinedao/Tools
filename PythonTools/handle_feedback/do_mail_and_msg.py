import win32com.client as win32
import shutil,os
from check_mail_phone import isPN,isEmail

xlApp = win32.gencache.EnsureDispatch('excel.application')
xlApp.Visible = False

# ------ 请改成要处理的文件名，放在脚本同目录下，（也可以放在其它目录，这里写完整路径）
filepath = os.getcwd()+'/用户反馈20150208.xlsx'
# ------

'''
只会处理操作问题，Bug反馈两种类型，excel中O列（处理人后面第二列）可以写自定义的回复，会根据用户所留的联系方式自动组织邮件或短信，
若O列（处理人后面第二列）是空的，且备注中子类型在回复模版中有数据的，就会按模版组织邮件或短信

除此之外，excel备注中写?会发送询问邮件或短信
注意在excel里面处理人留空，若不留空，则不会对行数据进行自动处理

'''

wb = xlApp.Workbooks.Open(filepath)
ws = wb.Worksheets("所有反馈")
# rowCountAllFeed = ws.Range('A65536').End(win32.constants.xlUp).Row
#从服务器上获取一份最新的模版到本地，本地若存在则覆盖
shutil.copyfile("//Share/用户反馈数据统计分类.xlsx",
                os.getcwd()+"/用户反馈数据统计分类.xlsx")
wbTemplate = xlApp.Workbooks.Open(os.getcwd()+"/用户反馈数据统计分类.xlsx",ReadOnly=1) #只读打开
wsTemplate = wbTemplate.Worksheets("回复模版")
rowCountTemplate = wsTemplate.Range('A65536').End(win32.constants.xlUp).Row   #UsedRange.Rows.Count
rowCountTemplate2 = wsTemplate.UsedRange.Rows.Count
if rowCountTemplate < rowCountTemplate2: rowCountTemplate = rowCountTemplate2
rowCountAllFeed = ws.UsedRange.Rows.Count
rowCountAllFeed2 = ws.Range('A65536').End(win32.constants.xlUp).Row
if rowCountAllFeed < rowCountAllFeed2: rowCountAllFeed = rowCountAllFeed2
print(rowCountAllFeed)
dealCollections = ("操作问题","Bug反馈")
dictemail = {}
dictPN = {}


# 获取当前机器的登录用户名，用于处理人签名
loguser = str(os.popen('echo %username%').read())
loguser = loguser.strip('\n')
ccto = loguser + "@xxxx.xxx" #自动生成自己的邮箱地址以便抄送：登录用户名-域账户加邮箱后缀
print(ccto)

########################################################################################################
########################## 为避免误操作，请在确实需要发送邮件时再临时取消注释 ##########################
########################################################################################################
def mailbyoutlook(to,ccto,subject,messagebody):
    app= 'Outlook'
    olook = win32.gencache.EnsureDispatch("%s.Application" % app)
    mail=olook.CreateItem(win32.constants.olMailItem)
    mail.Recipients.Add(to)
    # mail.Recipients.Add(ccto)
    subj = mail.Subject = subject
    body = messagebody
    mail.Body = body
    ###mail.Send()
    print("send ok")
########################################################################################################

dictTemplateEmail = {} # 存储邮件模版的字典，key是excel第一列类别名
dictTemplateMsg = {} # 存储短信模版的字典，key是excel第一列类别名

# 把回复模版分别读入邮件字典和短信字典
for r in range(2,rowCountTemplate + 1):
    k = wsTemplate.Cells(r,1).Value
    v = wsTemplate.Cells(r,2).Value
    m = wsTemplate.Cells(r,3).Value
    if(k in dictTemplateEmail):
        print("Error! Duplicated collections!", k)
    else:
        dictTemplateEmail[k] = v
    if(k in dictTemplateMsg):
        print("Error! Duplicated collections!", k)
    else:
        dictTemplateMsg[k] = m

# 邮件正文
msgbody_header = '''您好，

您反馈的问题：'''

# 实现换行效果
linefeed = '''

'''

askdetails = '''

请您具体描述一下您的问题，在什么情况下出现此问题，详细操作步骤，问题屏幕截图等。

'''

msgbody_footer = '''

如还有问题，请联系xxx@xxxxx.xxx或QQ群xxxxxx ^_^ 谢谢。
'''




try:
    for r in range(1,rowCountAllFeed + 1):
        if(ws.Cells(r,13).Value == None):  # 只有处理人一栏为空时，才会处理
            # if(ws.Cells(r,1).Value in dealCollections):
            if(ws.Cells(r,5).Value != None):
                if isEmail(ws.Cells(r,5).Value):
                    dictemail[r]=ws.Cells(r,5).Value
                elif isPN(ws.Cells(r,5).Value):
                    dictPN[r] = ws.Cells(r,5).Value
            if(ws.Cells(r,4).Value != None):
                if isEmail(ws.Cells(r,4).Value):
                    dictemail[r]=ws.Cells(r,4).Value
                elif isPN(ws.Cells(r,4).Value):
                    dictPN[r] = ws.Cells(r,4).Value

            # if(ws.Cells(r,12).Value != None):
            if r in dictemail:
                if ws.Cells(r,15).Value != None: # and ws.Cells(r,13).Value == None: # 优先发送自定义回复内容，而不是模版，请在O列填写自定义回复
                    if(ws.Cells(r,12).Value==None):
                        subject = "客户端用户反馈"
                    else:subject = "客户端用户反馈--"+ws.Cells(r,12).Value
                    msgbody = msgbody_header + ws.Cells(r,2).Value + linefeed + ws.Cells(r,15).Value  + msgbody_footer
                    mailbyoutlook(dictemail[r],ccto,subject,msgbody)
                    ws.Cells(r,13).Value = loguser + "_mail"
                    print("\n\n=========\n\n")
                    print(dictemail[r])
                    print(msgbody)
                # 发送模版
                elif ws.Cells(r,1).Value in dealCollections and ws.Cells(r,12).Value in dictTemplateEmail and ws.Cells(r,2).Value != None and ws.Cells(r,12).Value != None :
                    subject = "客户端用户反馈--"+ws.Cells(r,12).Value
                    msgbody = msgbody_header + ws.Cells(r,2).Value + linefeed + dictTemplateEmail[ws.Cells(r,12).Value]
                    # 邮件发送完毕后，处理人一栏填上当前电脑用户的拼音和后缀_mail
                    mailbyoutlook(dictemail[r],ccto,subject,msgbody)
                    ws.Cells(r,13).Value = loguser + "_mail"
                    print("\n\n=========\n\n")
                    print(dictemail[r])
                    print(msgbody)
            elif r in dictPN:
                if ws.Cells(r,15).Value != None: #  and ws.Cells(r,13).Value == None: # 优先发送自定义回复内容，而不是模版，请在O列填写自定义回复
                    ws.Cells(r,14).Value = "亲，感谢您的反馈：" + ws.Cells(r,15).Value + "如还有问题，请联系xxxxxx@xxxxxx.xxx或QQ群xxxxxx ^_^"
                    ws.Cells(r,13).Value = loguser + "_msg"
                    print(len(ws.Cells(r,14).Value))
                    if len(ws.Cells(r,14).Value) > 130:
                        print("=======================================To Long================")
                # 模版
                elif ws.Cells(r,1).Value in dealCollections and ws.Cells(r,12).Value in dictTemplateMsg:
                    ws.Cells(r,14).Value = dictTemplateMsg[ws.Cells(r,12).Value]
                    ws.Cells(r,13).Value = loguser + "_msg"

            # 发送详情询问邮件
            # print(ws.Cells(r,13).Value,ws.Cells(r,12).Value, r in dictemail)
            if(ws.Cells(r,13).Value==None and ws.Cells(r,12).Value in ("？","?")):
                if isEmail(ws.Cells(r,4).Value): # or isEmail(ws.Cells(r,5).Value):
                    dictemail[r] = ws.Cells(r,4).Value
                elif isPN(ws.Cells(r,4).Value):
                    dictPN[r] = ws.Cells(r,4).Value
                elif isEmail(ws.Cells(r,5).Value):
                    dictemail[r] = ws.Cells(r,5).Value
                elif isPN(ws.Cells(r,5).Value):
                    dictPN[r] = ws.Cells(r,5).Value
                if r in dictemail:
                    subject = "客户端用户反馈"
                    msgbody = msgbody_header + ws.Cells(r,2).Value  + askdetails + msgbody_footer
                    mailbyoutlook(dictemail[r],ccto,subject,msgbody)
                    ws.Cells(r,13).Value = loguser + "_mail"
                elif r in dictPN:
                    ws.Cells(r,14).Value = "亲，感谢您的反馈：请您帮忙提供一下更详细更具体的信息到xxxxx@xxx.xxx或QQ群xxxxxx ^_^"
                    ws.Cells(r,13).Value = loguser + "_msg"
finally:
    wbTemplate.Close()
    wb.Save()
    wb.Close()
print("============= done =============")
os.startfile(filepath)