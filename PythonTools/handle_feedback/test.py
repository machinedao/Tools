import os
import string


filelist = []


for filename in os.listdir(os.getcwd()+"/resources"):
    filelist.append(filename)

for i in filelist:
    print(i)

print(len(filelist))

for i in string.ascii_lowercase:
    print(i,end='')
