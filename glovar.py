#!/usr/bin/python
#-*-coding:UTF-8-*-

#some global varibles
#database param for ds(data source),xf(xuefu)
db4ds = {"database":"zhijin_data","user":"myedu","password":"dbpassword","host":"172.16.40.4","port":"3433"}
db4xf = {"database":"education","user":"myedu","password":"dbpassword","host":"172.16.40.4","port":"3433"}
attchSvr = {"host":"172.16.33.67","port":22,"username":"root","password":"u070Px%9P12!"}
currentBatch = ""
currentBatchId = 0
styleCommon = None
styleAlert = None
styleCommonDate = None
styleAlertDate = None