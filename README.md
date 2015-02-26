# Tools
some tools for daily work

handle_feedback
提交以前用来辅助回复用户投诉反馈的一些脚本，win7，office2010
主要有：get_excel_from_db.py
每天从数据库中直接获取一天的数据并套用excel模版生成件
filter.py 根据关键词智能分类常见投诉反馈
do_mail_and_msg.py 用来根据模版调用outlook自动发送邮件或组织短信内容
summary.py 和 weeklydetails.py用于每周汇总报告

PythonTools/Misc/try_mongodb.py  连接和查询MongoDb的一个简单例子

BashTools/DailySignin  测试每日签到，需要把服务器的时间改后一天，然后再签到，写了个bash + python 脚本来实现，python 负责调用签到接口，依赖bash所传的token打开制定URL，而bash除了把token传进python，还负责循环次数，即需要自动签到几天，./nextday_signin.sh token times 这样使用即可
