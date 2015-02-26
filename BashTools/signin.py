import urllib2,os,sys

# （注意第 0 个参数是脚本名本身）

thedate = sys.argv[1]
_token = sys.argv[2]

url1 = "http://192.168.xxx.xxx:8090/mobile/api2.do?action=xxxxSignin&signinDate=%s&token=%s" % (thedate,_token)
url2 = '''&returnType=json&deviceType=Android&xxx'''

url = url1 + url2

data=urllib2.urlopen(url)

print(thedate,"python done")

