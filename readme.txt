#ttAuditor
#version: 4.5
#author: 孙蕾

New Feature in 4.5:
	1.派单策略改为：备注 包含 ‘铁通转网’ && 环节追踪的最后一个环节 包含 ‘前台预约’ && 状态 包含 ‘成功’
	2.减少等待时间，优化性能
	3.[debugMode]不同的值有不同的行为，当且仅当值为‘0’时真正派单

New Feature in 4.0:
	1.fix firefox有时进入主页失败的bug
	*2.实现自动化派单。config中新增[assignMode],使用方法如下
		2.1 [useMode]:值为[fullAuto]时全自动运行，需要在[fullAuto]->[assignCount]设置次数
		（即为所有可以派单的员工每人派单的次数）
		2.2 之前版本的默认使用的半自动模式，可以将[useMode]设为[halfAuto]以开启，其下各个参数的含义相同
	3.config中新增[debugMode],正常使用时请设为0，测试时设为1（程序会运行但不会真正派单）

New Feature in 3.0:
	1.工单首页可自动翻页
	2.预约日期自动设置为下个月底
	3.寻找控件的成功率提升

-------------------------------
1.按照之前的要求搭好环境

2.安装firefox, 注意，必须将其安装在D:\Program Files (x86)\Mozilla Firefox目录下。
**安装完后请确保该目录下有firefox.exe，并可以使用

3.配置tt.config
3.0 mode设置运行方式，设为'amazingTT'为实时UI运行（**建议初次使用时以此实验），设为'standardTT'则为后台运行
3.1 assignWorker设置此次需要派单的员工名称
3.2 assignCount设置此次此员工需要派单的数量


To Do:
1.打包为exe文件