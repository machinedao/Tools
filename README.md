# Tools
some tools for daily work

PythonTools/Misc/try_mongodb.py  连接和查询MongoDb的一个简单例子

BashTools/DailySignin  测试每日签到，需要把服务器的时间改后一天，然后再签到，写了个bash + python 脚本来实现，python 负责调用签到接口，依赖bash所传的token打开制定URL，而bash除了把token传进python，还负责循环次数，即需要自动签到几天，./nextday_signin.sh token times 这样使用即可
