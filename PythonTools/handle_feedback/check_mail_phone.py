import re

# 判断是否是邮箱地址
def isEmail(a):
    a = str(a)
    if a == None:
        return False
    elif a.find('@') > 0:
        invalidEmail = ["_user.com",]  # QQ、新浪等第三方登录的会产生含_user.com非真实邮箱地址
        for i in invalidEmail:
            if a.find(i) > 0:
                return False
        return True
    else:
        return False

#判断是否手机号
def isPN(a):
    # a = str(a)
    if a == None:
        return False
    elif len(str(a)) > 12: # 从excel中过来的手机号多一个字符，故写成12
        # print(a,False)
        return False
    elif re.match("1[0-9]{10}",a):
        # print(a,True)
        return True
    else:
        # print(a,False)
        return False