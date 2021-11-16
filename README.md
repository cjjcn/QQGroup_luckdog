# 使用说明

此脚本采用chromedriver驱动，模拟打开qun.qq.com后经过一次iframe认证后进入群成员列表界面后再进行数据爬取，已经过chrome(95.0.4638.69)测试正常。

在`main.py`中按照需要改好爬取的QQ群号，并确认chromedrive路径正确配置即可运行，运行后爬取的信息会以xls的格式存储在同一目录下。

爬取出信息后即可使用`luckdog.py`，修改好xls文件名后运行即可进行群成员抽取，可用于课堂随机点名，群内随机抽奖等活动。

# 联系方式

[jiajchencn@gmail.com](mailto:jiajchencn@gmail.com)