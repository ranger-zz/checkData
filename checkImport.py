#!/usr/bin/python
#-*-coding:UTF-8-*-
import xlrd,xlwt,os,time,datetime,psycopg2,sys,math
import paramiko
import glovar

glovar.styleCommon = xlwt.easyxf('font: name Times New Roman, color-index black')
glovar.styleAlert = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
glovar.styleCommonDate = xlwt.easyxf('font: name Times New Roman, color-index black',num_format_str='M/D/YY')
glovar.styleAlertDate = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='M/D/YY')

#select output type
def inputExcelFileName():
    fileName = "";
    fileName = raw_input("input the file name to process in directory tmp: ")
    fileName.strip()
    if os.path.isfile("tmp/" + fileName):
    	pass
    else:
    	print("file doesn't exist!")
    	fileName = ""
    contents = getContentFromXls("tmp/" + fileName)
    if contents == None:
    	print("file is damaged! connot read!")
    	fileName = ""
    	pass
    return fileName
#select output type
def inputImportStartLine():
    startLine = -1;
    try:
	    startLine = int(raw_input("input the start line number(starting from number 1): "))
    except Exception, e:
    	startLine = -1
    else:
    	startLine = startLine - 1
    finally:
    	pass
    return startLine

#select function
def selectFromDB(dbParam,sql):
    #print(dbParam["database"] + "/" + dbParam["user"] + "/" + dbParam["password"] + "/" + dbParam["host"])
    conn = psycopg2.connect(database=dbParam["database"],user=dbParam["user"],password=dbParam["password"],host=dbParam["host"],port=dbParam["port"])
    cur = conn.cursor()
    cur.execute(sql)
    rows = cur.fetchall()
    conn.close()  
    return rows

#update or delete function
def executeFromDB(dbParam,sqls):
    conn = psycopg2.connect(database=dbParam["database"],user=dbParam["user"],password=dbParam["password"],host=dbParam["host"],port=dbParam["port"])
    cur = conn.cursor()
    try:
	    for sql in sqls:
	    	#print(sql)
	    	cur.execute(sql)
	    conn.commit()
    except Exception, e:
    	print(str(e))
    	conn.rollback()
    else:
    	pass
    finally:
    	pass
    conn.close()

#get xls file content
def getContentFromXls(fileName):
	try:
		content = xlrd.open_workbook(fileName)
	except Exception, e:
		print(str(e))
		return None		
	else:
		return content
	finally:
		pass
	pass

#translate list to string guide by columnName and type
def transFromList2Str(colDef,colType,tList):
	listStr = ""
	for i in xrange(0,len(colDef)):
		tp = colType[i]
		cl = tList[i]
		cn = colDef[i]
		listStr = listStr + cn + ":"
		if cl == None:
			listStr = listStr + ","
		else:
			if tp == 0:
				listStr = listStr + ","
			elif tp == 1:
				listStr = listStr + cl + ","
			elif tp == 2:
				listStr = listStr + str(cl) + ","
			elif tp == 3:
				listStr = listStr + cl.strftime("%Y-%m-%d") + ","
			elif tp == 4:
				if cl:
					listStr = listStr + "True" + ","
				else:
					listStr = listStr + "False" + ","
					pass
			else:
				listStr = listStr + ","
		pass
	pass
	listStr = listStr[:len(listStr) - 1]
	return listStr

#sort list by column
#colNum: sort field
#tList: list to sort
#direction: 0:asc 1:desc
def sortListByCol(colNum,tList,direction):
	listLength = len(tList)
	for i in range(listLength)[::-1]:
		for j in range(i):
			id1 = tList[j][colNum]
			id2 = tList[j+1][colNum]
			if direction == 0:
				if id1>id2:
					tList[j],tList[j+1] = tList[j+1],tList[j]
		pass
	pass

#get a null row
#colType is a list to define row's fields type
def getRowWithNone(colType):
	tmprow = []
	for i in range(len(colType)):
		if colType[i]==1:
			tmprow.append("")
		else:
			tmprow.append(None)
	return tmprow

#locate the record in list
#return -1 if not in the list
#this is a simplyfied alogrithm, the pre-condition is the srcList had been sorted
#direction indicats the direction of srcList sorted
def locateInSortedList(value,colNum,srcList,direction):
	index = -1
	for i in range(len(srcList)):
		tmpValue = srcList[i][colNum]
		if value == tmpValue:
			index = i
			break
		elif value > tmpValue:
			if direction == 0:
				continue
			else:
				break
		else:
			if direction == 0:
				break
			else:
				continue
		pass
	pass
	return index

def findIndexInList(value,tList):
	index = -1
	for i in range(len(tList)):
		if value==tList[i]:
			index = i
			break
	return index

def progressbar(cur, total):
	percent = '{:.2f}'.format(float(cur) / float(total))
	sys.stdout.write('\r')
	sys.stdout.write("[%-50s] %s" % ('=' * int(math.floor(cur * 50 / total)),percent))
	sys.stdout.flush()

#####################################################################################################
#										Write List into Excel fileName								#
#####################################################################################################
#1st column define the style,0-common,1-alert
def write2Excel(tList,wsName,dstFile,columns):
	wb = xlwt.Workbook(encoding='utf-8')
	ws = wb.add_sheet(wsName)
	#write example
	#ws.write(0, 0, 1234.56, style0)
	#ws.write(1, 0, datetime.now(), style1)
	#ws.write(2, 0, 1)
	#ws.write(2, 2, xlwt.Formula("A3+B3"))

	#1st row for column name
	for i in range(len(columns)):
		ws.write(0,i,columns[i],glovar.styleCommon)
		pass
	#write result list
	for i in range(len(tList)):
		tmprow = tList[i]
		style = tmprow[0]
		for j in range(len(columns)):
			value = tmprow[j + 1]
			#print(type(value))
			#print(value)
			if style==0:
				if value == None:
					ws.write(i + 1,j,"",glovar.styleCommon)
				elif isinstance(value, datetime.datetime) or isinstance(value, datetime.date):
					ws.write(i + 1,j,value,glovar.styleCommonDate)
				else:
					ws.write(i + 1,j,value,glovar.styleCommon)
			elif style==1:
				if value == None:
					ws.write(i + 1,j,"",glovar.styleAlert)
				elif isinstance(value, datetime.datetime) or isinstance(value, datetime.date):
					ws.write(i + 1,j,value,glovar.styleAlertDate)
				else:
					ws.write(i + 1,j,value,glovar.styleAlert)				
		pass

	if os.path.isfile(dstFile):
		os.remove(dstFile)
		pass

	wb.save(dstFile)
#####################################################################################################
#										Write List into Excel fileName								#
#####################################################################################################
#####################################################################################################
#								if a string varible is null or empty								#
#####################################################################################################
# if a Nonetype varible return 1
# if a string with empty chars return 2
def ifValueEmpty(varible):
	ret = 0
	if varible==None:
		ret = 1
	elif len(varible)==0:
		ret = 2
	return ret
#####################################################################################################
#								if a string varible is null or empty								#
#####################################################################################################
#####################################################################################################
#								Intialize Program Constants and Virable								#
#####################################################################################################
def initialize():
	#initialize the glovar.currentBatch
	sql = "select batch_name,batch_id from batch where is_current='1'"
	rows = selectFromDB(glovar.db4xf,sql)
	glovar.currentBatch = rows[0][0]
	glovar.currentBatchId = rows[0][1]

	reload(sys)  
	sys.setdefaultencoding('utf8')   
	pass
#####################################################################################################
#								Intialize Program Constants and Virable								#
#####################################################################################################
#####################################################################################################
#								Check Import Unmatched Students List								#
#####################################################################################################
def checkImport():
	#identity cards in Excel
	ids = []
	#Excel rows, initialized by reading file, appended by comparing with ds and xf
	#thus, every row in rows will expend to three times during the comparation
	#including xlsrows,dsrows and xfrows
	#dsrows2C has the same row number as xlsrows, and every row has the same identity_card as in the xlsrows. 
	#so, commonly, some rows in dsrows2C only have 2 values: source(named ds),identity_card. and other columns would be none
	#the same to xfrows2C
	xlsrows = []
	dsrows2C = []
	xfrows2C = []
	#memorows are used to record errors in xlsrows
	memoRows = []
	columnName = []
	columnType = []

	rows = []
	srcFile = ""
	dstFile = ""

	#input file name to check
	fileName = inputExcelFileName()
	while len(fileName) == 0:
		fileName = inputExcelFileName()
	srcFile = "tmp/" + fileName
	dstFile = "result/self-" + fileName
	dstFileError = "result/error-" + fileName 
	dstFileImport = "result/import-" + fileName
	print(srcFile + "--->>>" + dstFile)

	checkUnmatchFile(srcFile,ids,xlsrows,dsrows2C,xfrows2C,memoRows,columnName,columnType)
	#print(str(len(xlsrows)) + "/" + str(len(dsrows2C)) + "/" + str(len(xfrows2C)) + "/" + str(len(memoRows)))
	verifyStudentField(xlsrows,dsrows2C,xfrows2C,memoRows,1)
	#print(str(len(xlsrows)) + "/" + str(len(dsrows2C)) + "/" + str(len(xfrows2C)) + "/" + str(len(memoRows)))

	colNum = len(columnType)
	results = []
	for i in range(len(xlsrows)):
		tmprow = []
		tmprow.append(0)
		for j in range(colNum):
			tmprow.append(xlsrows[i][j])
		results.append(tmprow)
		tmprow = []
		tmprow.append(0)
		for j in range(colNum):
			tmprow.append(dsrows2C[i][j])
		results.append(tmprow)
		tmprow = []
		tmprow.append(0)
		for j in range(colNum):
			tmprow.append(xfrows2C[i][j])
		results.append(tmprow)
		tmprow = []
		tmprow.append(1)
		for j in range(colNum):
			tmprow.append(memoRows[i][j])
		results.append(tmprow)
	write2Excel(results,"result",dstFile,columnName)

	resultsError = []
	resultsImport = []
	for i in range(len(xlsrows)):
		tmprow = []
		memoRows[i][0] = ""
		style = any(memoRows[i])
		memoRows[i][0] = "memo"
		#print(transFromList2Str(columnName,columnType,memoRows[i]))
		if style == False:#empty list
			tmprow.append(0)
		else:
			tmprow.append(1)
		for j in range(colNum):
			msg = memoRows[i][j]
			if ifValueEmpty(msg)>0:#empty value
				tmprow.append(xlsrows[i][j])
			else:
				tmprow.append(xlsrows[i][j] + "\r\n" + memoRows[i][j])
		if style:
			resultsError.append(tmprow)
		else:
			resultsImport.append(tmprow)
		#results1.append(tmprow)
	write2Excel(resultsError,"error",dstFileError,columnName)
	write2Excel(resultsImport,"result",dstFileImport,columnName)
	mainMenu()	

def checkUnmatchFile(srcFile,ids,xlsrows,dsrows2C,xfrows2C,memoRows,columnName,columnType):
	#column definition
	columnName1 = ["source","enroll_batch","study_center_name","enroll_school","enroll_arrangement","major","student_name",
	"identity_card","enrollment_study_center","enrollment_people","student_source","project_type","enroll_date","manage_type",
	"manage_study_center","service_center","department","refer_name","refer_batch","refer_identity_card","input_date"]
	#0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
	columnType1 = [1,1,1,1,1,1,1,1,1,1,1,1,3,1,1,1,1,1,1,1,3]

	for i in range(len(columnName1)):
		columnName.append(columnName1[i])
		columnType.append(columnType1[i])

	dsrows = []
	xfrows = []

	#input file name to check
	startLine = inputImportStartLine()
	while startLine < 0:
		startLine = inputImportStartLine()

	content = getContentFromXls(srcFile)
	worksheet = content.sheets()[0]
	#initialize ids 
	ids1 = worksheet.col_values(6)
	j = 0;
	for i in range(startLine,len(ids1)):
		ids.append(ids1[i])
		j = j + 1
	#get rows from xls
	total = worksheet.nrows - startLine
	print("analyze the file... total records sums to " + str(total))
	for i in range(startLine,worksheet.nrows):
		progressbar(i - startLine,total)
		tmprow = []
		tmprow.append("xls")
		row = worksheet.row_values(i)
		for j in range(len(columnType) - 1):
			ctype = 0
			if j >= worksheet.ncols:
				value = None
			else:
				value = row[j]
				ctype = worksheet.cell(i,j).ctype
			if ctype==0:
				if j==11 or j==19:
					value=None
				pass
			elif ctype==1:
				value = value.strip()
				if j==11 or j==19:
					if ifValueEmpty(value)==0:
						#print(value)
						value = datetime.datetime.strptime(value,"%Y-%m-%d")
					else:
						value = None
			elif ctype==2:
				if j==0 or j==6 or j==17 or j==18:
					value = str(int(value))
			elif ctype==3:
				value = xlrd.xldate.xldate_as_datetime(value,0)
			elif ctype==4:
				pass
			else:
				pass
			tmprow.append(value)
		#print(transFromList2Str(columnName,columnType,tmprow))
		xlsrows.append(tmprow)
	sortListByCol(7,xlsrows,0)
	print("")

	#get record from ds
	print("get the records from data source... ")
	sql = "select enroll_batch,study_center_name,enroll_school,enroll_arrangement,major,student_name,identity_card,enroll_date "
	sql = sql + "from crawler_student where enroll_batch='" + glovar.currentBatch + "' and status=1 "
	sql = sql + "and identity_card in("
	for i in xrange(0,len(ids)):
		sql = sql + "'" + ids[i] + "',"
	sql = sql[:len(sql) - 1] + ") order by identity_card"
	#print(sql)
	rows = selectFromDB(glovar.db4ds,sql)
	total = len(rows)
	print("total records in ds sums to " + str(total))
	i = 0;
	for row in rows:
		progressbar(i,total)
		i = i + 1
		tmprow = []
		tmprow.append("ds")
		for j in xrange(0,len(row)):
			#6 columns ahead are in sequence of the colunmdefinition
			value = row[j]
			if j < 7:
				if j>0 and j<6:
					if value!=None:
						value = value
					else:
						value = ""
				tmprow.append(value)
			else:
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append(value)
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append("")
				tmprow.append(None)
		#print(transFromList2Str(columnName,columnType,tmprow))
		dsrows.append(tmprow)
	print("")

	#get record from xf
	print("get the records from xuefu... ")
	sql = "select s.enroll_batch,se.study_center_name,se.enroll_school,se.enroll_arrangement,se.major,s.student_name,s.identity_card,"
	sql = sql + "s.enrollment_study_center,se.enrollment_people,se.student_source,s.project_type,se.enroll_date,se.manage_type,"
	sql = sql + "se.manage_study_center,se.service_center,se.department,s1.student_name,s1.enroll_batch,s1.identity_card,se.input_date "
	sql = sql + "from student as s left join student_enroll as se on s.student_id=se.student_id "
	sql = sql + "left join student as s1 on s.refer_student_id=s1.student_id " + "where "
	sql = sql + "s.enroll_batch='" + glovar.currentBatch + "' "
	sql = sql + "and (s.is_delete_flag is null or s.is_delete_flag='1') "
	sql = sql + "and s.identity_card in("
	for i in xrange(0,len(ids)):
		sql = sql + "'" + ids[i] + "',"
	sql = sql[:len(sql) - 1] + ") order by s.identity_card"
	#print(sql)
	rows = selectFromDB(glovar.db4xf,sql)
	total = len(rows)
	print("total records in xf sums to " + str(total))
	i = 0;
	for row in rows:
		progressbar(i,total)
		i = i + 1
		tmprow = []
		tmprow.append("xf")
		for j in xrange(0,len(row)):
			#6 columns ahead are in sequence of the colunmdefinition
			'''
			if j == 16:
				#refer_student_id
				if row[j] != None or row[j] == "":
					sql1 = "select student_name,enroll_batch,identity_card from student where student_id=" + str(row[j])
					rows1 = selectFromDB(glovar.db4xf,sql1)
					tmprow.append(rows1[0][0])
					tmprow.append(rows1[0][1])
					tmprow.append(rows1[0][2])				
				else:
					tmprow.append("")
					tmprow.append("")
					tmprow.append("")
			else:
			'''
			value = row[j]
			if (j>0 and j<6) or (j>6 and j<11) or (j>11 and j<19):
				if value!=None:
					value = value
				else :
					value = ""
			tmprow.append(value)
		#print(transFromList2Str(columnName,columnType,tmprow))
		xfrows.append(tmprow)
	print("")

	#compose xls-ds-xf into one list
	#if dsrows or xfrows do not have the record in xlsrows, append a null row
	total = len(xlsrows)
	print("compare xls between ds and xf... total records:" + str(total))
	#memoRows = []
	for i in range(len(xlsrows)):
		progressbar(i,total)
		xlsrow = xlsrows[i]
		identity_card = xlsrow[7]
		dsrow = []
		xfrow = []
		memorow = getRowWithNone(columnType)
		memorow[0] = "memo"
		memoRows.append(memorow)
		index = locateInSortedList(identity_card,7,dsrows,0)
		if index < 0:
			dsrow = getRowWithNone(columnType)
			dsrow[0] = "ds"
		else:
			dsrow = dsrows[index]
		dsrows2C.append(dsrow)
		index = locateInSortedList(identity_card,7,xfrows,0)
		if index < 0:
			xfrow = getRowWithNone(columnType)
			xfrow[0] = "xf"
		else:
			xfrow = xfrows[index]
		xfrows2C.append(xfrow)
	print("")
#####################################################################################################
#								Check Import Unmatched Students List								#
#####################################################################################################
#####################################################################################################
#								Verify every fields of student infomation							#
#####################################################################################################
#checkType = 1 means: fields need to compare with xf records, =0 means do not compare with xf records
def verifyStudentField(xlsrows,dsrows2C,xfrows2C,memoRows,checkType):
	#define some constant
	#study_center_name in columnName index
	indexofSCN = 2
	#enroll_school in columnName index
	indexofES = 3
	#major in columnName index
	indexofEA = 4
	#enroll_arrangement in columnName index
	indexofMA = 5
	#student_name in columnName index
	indexofSN = 6
	#identity_card in columnName index
	indexofID = 7
	#enrollment_study_center in columnName index
	indexofESC = 8
	#enrollment_people in columnName index
	indexofEP = 9
	#student_sourc in columnName index
	indexofSS = 10
	#project_type in columnName index
	indexofPT = 11
	#manage_type in columnName index
	indexofMT = 13
	#manage_study_center in columnName index
	indexofMSC = 14
	#refer_name in columnName index
	indexofRN = 17
	#refer_batch in columnName index
	indexofRB = 18
	#refer_identity_card in columnName index
	indexofRIC = 19
	#direct enrollment student source
	ssForDE = ["慧生活","活动营销","兼职直招","金百泽企业","老带新","美的集团","亲属报读","上门咨询","团工委","网络营销","新带新","亦庄工会","员工报读",
	"员工直招","圆梦企业","圆梦社招","圆梦行业","专科续本","总部介绍","网络营销-新带新","网络营销-老带新","网络营销-直招","企业直招"]
	for i in range(len(ssForDE)):
		ss = ssForDE[i]
		ssForDE[i] = ss
		pass
	#manage type
	mtForCom = "中心直属管理"
	mtForChn = "咨询点管理"
	#student_source for "customer bring customer"
	ssForCBC = ["老带新","新带新","网络营销-新带新","网络营销-老带新"]

	#check every record's, verify every fields(except enroll_batch,enrollment_study_center,enroll_date,service_center,department,input_date)
	sql = "select dictionary_name from data_dictionary where parent_dictionary_id=48"
	recs = selectFromDB(glovar.db4xf,sql)
	#pts means project type list
	pts = []
	for i in range(len(recs)):
		rec = recs[i]
		pts.append(rec[0])
		pass
	total = len(xlsrows)
	print("verifing fields of every record... total records:" + str(total))
	for i in range(0,total):
		progressbar(i,total)
		row = xlsrows[i]
		
		#check if exist in xf
		if checkType == 1:
			xfrows2C[i][0] = ""
			#print(transFromList2Str(columnName,columnType,xfrows2C[i]))
			if any(xfrows2C[i])==False:
				xfrows2C[i][0] = "xf"
				memoRows[i][1] = "该学员在学服中没有记录"
				continue
			xfrows2C[i][0] = "xf"
			ep_xf = xfrows2C[i][indexofEP]
			ss_xf = xfrows2C[i][indexofSS]
			if ifValueEmpty(ep_xf)==0:#no a null
				memoRows[i][1] = "该学员已经完成导入匹配"
				if ifValueEmpty(ss_xf)>0:
					memoRows[i][1] = memoRows[i][1] + "/Null Student Source" 
				continue
		#check project type
		pt = row[indexofPT]

		if (checkType==0 and ifValueEmpty(pt)==0) or checkType==1:
			count = pts.count(pt)
			if count == 0:
				memoRows[i][indexofPT] = "项目类型不存在"

		#check enrollment_people,student_source
		ep = row[indexofEP]
		ep_xf = xfrows2C[i][indexofEP]
		if ep_xf!=None and checkType==1:
			if len(ep_xf) > 0:
				if ep != ep_xf:
					memoRows[i][indexofEP] = "招生人员与学服系统记录不符，学服记录为：" + ep_xf
				else:
					memoRows[i][indexofEP] = "该学员已经完成导入匹配"
		if (checkType==0 and ifValueEmpty(ep)==0) or checkType==1:
			sql = "select user_id from \"user\" as u where u.user_name='" + ep + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			userId = -1
			if len(recs) == 0:
				userId = -1
			else:
				userId = recs[0][0]
			if userId == -1:
				memoRows[i][indexofEP]="招生人员姓名有误"
			ss = row[indexofSS]
			count = ssForDE.count(ss)
			if count == 0:
				#not a direct enrollment student or wrong student_source
				if userId > 0:
					sql = "select c.audited_status,cp.audited_status,c.docking_people_id,u.user_name "
					sql = sql + "from channel as c left join channel_protocol as cp "
					sql = sql + "on c.channel_id=cp.channel_id "
					sql = sql + "left join \"user\" as u on c.docking_people_id=u.user_id "
					sql = sql + "where c.channel_name='" + ss + "' "
					sql = sql + "and (enroll_begin_date<='" + glovar.currentBatch + "' or enroll_end_date>='" + glovar.currentBatch + "') "
					recs = selectFromDB(glovar.db4xf,sql)
					if len(recs) == 0:
						memoRows[i][indexofSS] = "招生来源不存在"
					else:
						audit1 = recs[0][0]
						audit2 = recs[0][1]
						dpi = recs[0][2]
						dpn = recs[0][3]
						if audit1 != "2":
							memoRows[i][indexofSS] = "咨询点资质审核未通过"
						elif audit2 != "2":
							memoRows[i][indexofSS] = "咨询点协议审核未通过"
							pass
						else:
							if userId != dpi:
								memoRows[i][indexofSS] = "咨询点与招生人员不匹配，系统记录的咨询点对接人为" + dpn
								pass
				else:
					pass
			else:
				pass

		#check student_name and identity_card
		sn = row[indexofSN]
		idc = row[indexofID]
		if (checkType==0 and ifValueEmpty(sn)==0) or checkType==1:
			sql = "select s.student_name,u.user_name from student as s left join \"user\" as u on s.user_id=u.user_id "
			sql = sql + "where s.identity_card='" + idc + "' and s.is_delete_flag is null "
			sql = sql + "and s.enroll_batch='" + glovar.currentBatch + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofID] = "当前批次无此学员"
			else:
				rec = recs[0]
				snfdb = rec[0]
				if checkType==1:
					if snfdb != sn:
						memoRows[i][indexofSN] = "身份证与学员姓名不匹配，系统记录学员姓名为" + snfdb
					if rec[1] == None:
						pass
					else:
						unfdb = recs[0][1]
						if unfdb != ep and unfdb != ss and checkType==1:
							memoRows[i][indexofID] = "此学员已被其他老师或咨询点登记：" + unfdb

		#check studycenter,enroll_school,major and arrangement
		scn = row[indexofSCN]
		es = row[indexofES]
		major = row[indexofMA]
		ea = row[indexofEA]
		if ea=="高起专":
			row[indexofEA] = "专科"
			ea = "专科"

		scn_xf = xfrows2C[i][indexofSCN]
		es_xf = xfrows2C[i][indexofES]
		major_xf = xfrows2C[i][indexofMA]
		ea_xf = xfrows2C[i][indexofEA]

		scn_ds = dsrows2C[i][indexofSCN]
		es_ds = dsrows2C[i][indexofES]
		major_ds = dsrows2C[i][indexofMA]
		ea_ds = dsrows2C[i][indexofEA]

		veri4Fields = True
		if scn!=scn_xf and checkType==1:
			memoRows[i][indexofSCN] = "授权学习中心与学服系统记录不符，学服记录为：" + scn_xf + "/" + scn_ds
			veri4Fields = False
		if es!=es_xf and checkType==1:
			memoRows[i][indexofES] = "报名院校与学服系统记录不符，学服记录为：" + es_xf + "/" + es_ds
			veri4Fields = False
		if major!=major_xf and checkType==1:
			memoRows[i][indexofMA] = "报名专业与学服系统记录不符，学服记录为：" + major_xf + "/" + major_ds
			veri4Fields = False
		if ea!=ea_xf and checkType==1:
			memoRows[i][indexofEA] = "报名层次与学服系统记录不符，学服记录为：" + ea_xf + "/" + ea_ds
			veri4Fields = False
		if checkType==0:
			if ifValueEmpty(scn)>0 and ifValueEmpty(es)>0 and ifValueEmpty(major)>0 and ifValueEmpty(ea)>0:
				veri4Fields = False
			elif ifValueEmpty(scn)==0 and ifValueEmpty(es)==0 and ifValueEmpty(major)==0 and ifValueEmpty(ea)==0:
				pass	
			else:
				memoRows[i][indexofSCN] = "授权学习中心,报名院校,报名专业,报名层次必须同时为空，或者同时赋值"
				veri4Fields = False

		if veri4Fields==True:
			sql = "select organization_id from organization where organization_level='3' and organization_name='" + scn + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			scnid = -1
			esid = -1
			majorid = -1
			eaid = -1
			if len(recs) == 0:
				memoRows[i][indexofSCN] = "授权学习中心名称有误"
			else:
				scnid = recs[0][0]
			sql = "select college_id from college where college_name='" + es + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofES] = "院校名称有误"
			else:
				esid = recs[0][0]
			sql = "select major_id,college_id from major where major_name='" + major + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofMA] = "专业名称有误"
			else:
				majorid = recs[0][0]
				hit = 0
				for rec in recs:
					esid1 = rec[1]
					if esid1 != esid:
						memoRows[i][indexofMA] = "该专业非院校授权专业"
						majorid = -1
					else:
						memoRows[i][indexofMA] = ""
						majorid = rec[0]
						hit = 1
					if hit == 1:
						break
			sql = "select dictionary_id from data_dictionary where parent_dictionary_id=44 and dictionary_name='" + ea + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofEA] = "层次名称有误"
			else:
				eaid = recs[0][0]
				pass
			if scnid == -1 or esid == -1 or majorid == -1 or eaid == -1:
				pass
			else:
				sql = "select teaching_plan_id from teaching_plan where organization_id=" + str(scnid) 
				sql = sql + " and batch_id=" + str(glovar.currentBatchId)
				sql = sql + " and organization_id=" + str(scnid)
				sql = sql + " and college_id=" + str(esid)
				sql = sql + " and major_id=" + str(majorid)
				sql = sql + " and dictionary_id=" + str(eaid)
				recs = selectFromDB(glovar.db4xf,sql)
				if len(recs) == 0:
					memoRows[i][indexofSCN] = "当前院校、批次或层次未授权给当前学习中心"
					pass

		#check manage_type and manage_study_center
		mt = row[indexofMT]
		msc = row[indexofMSC]
		if (checkType==0 and ifValueEmpty(mt)==0) or checkType==1:
			if mt != mtForCom and mt != mtForChn:
				memoRows[i][indexofMT] = "学生分配管理方式有误"
			elif mt == mtForCom:
				if len(msc) == 0:
					memoRows[i][indexofMSC] = "中心直属管理下，管理学习中心不能为空"
					pass
				else:
					sql = "select organization_id from organization where organization_level='3' and organization_name='" + msc + "'"
					recs = selectFromDB(glovar.db4xf,sql)
					if len(recs) == 0:
						memoRows[i][indexofMSC] = "管理学习中心名称有误"
					else:
						msc_xf = xfrows2C[i][indexofMSC]
					#if msc!=msc_xf:
						#memoRows[i][indexofMSC] = "管理学习中心与学服记录不符，学服记录为：" + msc_xf
		if checkType==0 and ifValueEmpty(msc)==0:
			sql = "select organization_id from organization where organization_level='3' and organization_name='" + msc + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofMSC] = "管理学习中心名称有误"

		#check enroll_study_center
		esc = row[indexofESC]
		if ifValueEmpty(esc)==0:#enroll_study_center is not null
			sql = "select organization_id from organization where organization_level='3' and organization_name='" + esc + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofMSC] = "招生学习中心名称有误"
			
		#esc_xf = xfrows2C[i][indexofESC]
		#if esc!=esc_xf:
			#memoRows[i][indexofESC] = "招生学习中心与学服记录不符，学服记录为：" + esc_xf
		#check refer_name,refer_identity_card and refer_batch
		count = ssForCBC.count(ss)
		if count > 0:
			rn = row[indexofRN]
			rb = row[indexofRB]
			ric = row[indexofRIC]
			sql = "select enroll_batch,student_name from student where identity_card='" + ric + "'"
			recs = selectFromDB(glovar.db4xf,sql)
			if len(recs) == 0:
				memoRows[i][indexofRIC] = "推荐人不在学员记录里"
			else:
				rnfdb = recs[0][1]
				if rnfdb != rn and checkType==1:
					memoRows[i][indexofRN] = "推荐人姓名有误，系统记录推荐人姓名为" + rnfdb
				memoValue = ""
				for rec in recs:
					rbfdb = rec[0]
					if rbfdb == rb:
						memoValue = ""
						break
					else:
						if checkType==1:
							memoValue = memoValue + "推荐人批次有误，系统记录推荐人批次为" + rbfdb
					pass
				memoRows[i][indexofRB] = memoValue
	print("")
#####################################################################################################
#								Verify every fields of student infomation							#
#####################################################################################################
#####################################################################################################
#									Get Enrollment Data for Statistics 								#
#####################################################################################################
def getEnrollDataForStatistics():
	studycenterEnrollData = []
	branchList = ["北京分公司","山东分公司","上海分公司","广东分公司","安徽分公司","四川分公司","西安分公司","河南分公司","长春分公司"]
	for i in range(len(branchList)):
		branchList[i] = branchList[i]
	studycenterList = []
	for i in range(studycenterList):
		studycenterList[i] = studycenterList[i]
	studycenterInBranch = [0,1,1,1,1,1,2,3,3,3,3,3,4,5,6,7,8]
	#get study center enrollment data all status group by study_center_name(not enroll_study_center)
	sql = "select study_center_name,count(se.study_center_name) from student as s "
	sql = sql + "left join student_enroll as se on s.student_id=se.student_id "
	sql = sql + "where s.is_delete_flag is null"
	sql = sql + " and s.enroll_batch='" + glovar.currentBatch + "'"
	sql = sql + " and (se.study_center_name is not null or se.study_center_name!=''"
	sql = sql + " group by se.study_center_name"
	recs = selectFromDB(glovar.db4xf,sql)
	print("get all enrollment data with conditions: glovar.currentBatch,is_delete_flag is null,study_center_name is not null")
	total = len(recs)
	i = 0
	for rec in recs:
		progressbar(i,total)
		i = i + 1
		srec = []
		scn = rec[0]
		scnNum = rec[1]
		index = findIndexInList(scn,studycenterList)
		indexOfBranch = studycenterInBranch[index]
		srec.append(branchList[indexOfBranch])
		srec.append(scn)
		srec.append(scnNum)
		studycenterEnrollData.append(srec)
	sortListByCol(0,studycenterEnrollData,0)
	print("")
	mainMenu()
#####################################################################################################
#									Get Enrollment Data for Statistics								#
#####################################################################################################
def sftp_upload(host,port,username,password,local,remote):
    sf = paramiko.Transport((host,port))
    sf.connect(username = username,password = password)
    sftp = paramiko.SFTPClient.from_transport(sf)
    try:
        if os.path.isdir(local):#判断本地参数是目录还是文件
            for f in os.listdir(local):#遍历本地目录
                sftp.put(os.path.join(local+f),os.path.join(remote+f))#上传目录中的文件
        else:
            sftp.put(local,remote)#上传文件
    except Exception,e:
        print('upload exception:',e)
    sf.close()

def sftp_download(host,port,username,password,local,remote):
    sf = paramiko.Transport((host,port))
    sf.connect(username = username,password = password)
    sftp = paramiko.SFTPClient.from_transport(sf)
    try:
        if os.path.isdir(local):#判断本地参数是目录还是文件
            for f in sftp.listdir(remote):#遍历远程目录
                 sftp.get(os.path.join(remote+f),os.path.join(local+f))#下载目录中文件
        else:
            sftp.get(remote,local)#下载文件
    except Exception,e:
        print('download exception:',e)
    sf.close()

def updateChannelAttach():
	channelName = raw_input("input the channel full name:")
	sql = "select channel_img,channel_id from channel where channel_name='" + channelName + "'"
	recs = selectFromDB(glovar.db4xf,sql)
	if len(recs)==0:
		print("channel name is wrong")
	else:
		imageFile = raw_input("input image file path:")
		if os.path.isfile (imageFile):
			pos = imageFile.rfind(".")
			postfix = imageFile[pos:]
			rec = recs[0]
			channelId = rec[1]
			if rec[0]!=None:
				channelImg = rec[0]
			else:
				now = time.time() * 1000
				channelImg = channelName + "_" +'{0:13.0f}'.format(now) + postfix
			sql = "update channel set channel_img='" + channelImg + "' where channel_id=" + str(channelId)
			sqls = []
			sqls.append(sql)
			executeFromDB(glovar.db4xf,sqls)
			remote = "/mnt/ali-nas/cnxfu/education/storage/file/channel/" + channelImg
			print(imageFile + ">>>" + remote)
			host = glovar.attchSvr["host"]
			port = glovar.attchSvr["port"]
			username = glovar.attchSvr["username"]
			password = glovar.attchSvr["password"]
			sftp_upload(host,port,username,password,imageFile,remote)
		else:
			print("wrong image path!")
	pass
#####################################################################################################
#							upload attachment for student or channel								#
#####################################################################################################
def updateAttachments():
	updType = raw_input("input the object to update(1-student,2-channel):")
	if updType==1:
		pass
	elif updType==2:
		updateChannelAttach()
	mainMenu()
	pass
#####################################################################################################
#							upload attachment for student or channel								#
#####################################################################################################
#####################################################################################################
#										update student information									#
#####################################################################################################
def updateStudent():
	fileName = inputExcelFileName()
	while len(fileName)==0:
		fileName = inputExcelFileName()
	
	srcFile = "tmp/" + fileName
	dstFileError = "result/error-" + fileName 
	dstFileUpdate = "result/update-" + fileName
	file_path = "result/sql-" + fileName
	file_path = file_path[:len(file_path) - 4]
	#simpleFields=["enrollment_study_center","project_type","enroll_date","service_center","department","input_date"]
	#complexFileds=["enroll_batch","study_center_name","enroll_school","major","enroll_arrangement","student_name","identity_card",
	#"enrollment_people","manage_type","manage_study_center","refer_name","refer_batch","refer_identity_card"]

	columnName = []
	columnType = []
	xlsrows = []
	dsrows2C = []
	xfrows2C = []
	memoRows = []
	ids = []

	checkUnmatchFile(srcFile,ids,xlsrows,dsrows2C,xfrows2C,memoRows,columnName,columnType)
	verifyStudentField(xlsrows,dsrows2C,xfrows2C,memoRows,0)
	colNum = len(columnType)
	
	resultsError = []
	resultsImport = []
	sqls = []

	print("updating student information...")
	for i in range(len(xlsrows)):
		progressbar(i,len(xlsrows))
		tmprow = []
		memoRows[i][0] = ""
		style = any(memoRows[i])
		memoRows[i][0] = "memo"
		#print(transFromList2Str(columnName,columnType,memoRows[i]))
		if style == False:#empty list
			tmprow.append(0)
		else:
			tmprow.append(1)
		for j in range(colNum):
			msg = memoRows[i][j]
			if ifValueEmpty(msg)>0:#empty value
				tmprow.append(xlsrows[i][j])
			else:
				tmprow.append(xlsrows[i][j] + "\r\n" + memoRows[i][j])
		idsi = tmprow[8]
		if style:
			resultsError.append(tmprow)
		else:
			resultsImport.append(tmprow)
			sql = "select student_id from student where identity_card='" + idsi + "' and is_delete_flag is null and "
			sql = sql + "enroll_batch='" + glovar.currentBatch + "'"
			#print (sql)
			recs = selectFromDB(glovar.db4xf,sql)
			stuID = recs[0][0]
			userID = -1
			#dOrC indicates the student belongs zhijin or channel, 0 means direct, 1 means from channel, -1 means donot modify
			dOrC = -1
			if ifValueEmpty(tmprow[10])==0:
				dOrC = 0
				sql = "select user_id from \"user\" where user_name='" + tmprow[10] + "'"
				recs = selectFromDB(glovar.db4xf,sql)
				userID = recs[0][0]
			if ifValueEmpty(tmprow[11])==0:
				dOrC = 1
				sql = "select user_id from \"user\" where user_name='" + tmprow[11] + "'"
				recs = []
				recs = selectFromDB(glovar.db4xf,sql)
				if len(recs)>0:
					userID = recs[0][0]
			rid = -1
			if ifValueEmpty(tmprow[18])==0:
				ric = tmprow[20]
				rb = tmprow[19]
				sql = "select student_id from student where identity_card='" + ric + "' and is_delete_flag is null"
				sql = sql + " and enroll_batch='" + rb + "'"
				recs = selectFromDB(glovar.db4xf,sql)
				rid = recs[0][0]
			sql = "update student_enroll set "
			if ifValueEmpty(tmprow[3])==0:
				sql = sql + "study_center_name='" + tmprow[3] + "',"
			if ifValueEmpty(tmprow[4])==0:
				sql = sql + "enroll_school='" + tmprow[4] + "',"
			if ifValueEmpty(tmprow[6])==0:
				sql = sql + "major='" + tmprow[6] + "',"
			if ifValueEmpty(tmprow[5])==0:
				sql = sql + "enroll_arrangement='" + tmprow[5] + "',"
			if ifValueEmpty(tmprow[10])==0:
				sql = sql + "enrollment_people='" + tmprow[10] + "',"
			if ifValueEmpty(tmprow[11])==0:
				sql = sql + "student_source='" + tmprow[11] + "',"
			if ifValueEmpty(tmprow[14])==0:
				sql = sql + "manage_type='" + tmprow[14] + "',"
			if ifValueEmpty(tmprow[15])==0:
				sql = sql + "manage_study_center='" + tmprow[15] + "',"
			if ifValueEmpty(tmprow[16])==0:
				sql = sql + "service_center='" + tmprow[16] + "',"
			if ifValueEmpty(tmprow[17])==0:
				sql = sql + "department='" + tmprow[17] + "',"
			if tmprow[13]!=None:
				sql = sql + "enroll_date='" + tmprow[13].strftime("%Y-%m-%d") + "',"
			if tmprow[21]!=None:
				sql = sql + "input_date='" + tmprow[21].strftime("%Y-%m-%d") + "',"
			sql = sql[:len(sql)-1] + " "
			count = sql.count("=")
			if count>0:
				sql = sql + "where student_id=" + str(stuID)
				sqls.append(sql)
			sql = "update student set "
			if ifValueEmpty(tmprow[2])==0:
				sql = sql + "enroll_batch='" + tmprow[2] + "',"
			if ifValueEmpty(tmprow[7])==0:
				sql = sql + "student_name='" + tmprow[7] + "',"
			if ifValueEmpty(tmprow[9])==0:
				sql = sql + "enrollment_study_center='" + tmprow[9] + "',"
			if ifValueEmpty(tmprow[12])==0:
				sql = sql + "project_type='" + tmprow[12] + "',"
			if userID>0:
				sql = sql + "user_id=" + str(userID) + ","
			if rid>0:
				sql = sql + ",refer_student_id=" + str(rid) + ","
			sql = sql[:len(sql) - 1] + " "
			count = sql.count("=")
			if count>0:
				sql = sql + " where student_id=" + str(stuID)
				sqls.append(sql)
			if dOrC>0:
				sql = "select standby2 from \"user\" where user_id=" + str(userID)
				recs = selectFromDB(glovar.db4xf,sql)
				branch = recs[0][0]
				sql = "update student set branch_company='" + branch + "'"
				sqls.append(sql)
			#executeFromDB(glovar.db4xf,sqls)
		#results1.append(tmprow)
	print("")
	write2Excel(resultsError,"error",dstFileError,columnName)
	write2Excel(resultsImport,"result",dstFileUpdate,columnName)

	fw = open(file_path,"w")
	sql = ""
	for i in range(len(sqls)):
		sql = sql + sqls[i] + ";\n"
	fw.write(sql)
	fw.close()    
	executeFromDB(glovar.db4xf,sqls)
	mainMenu()
	pass
#####################################################################################################
#									update student information										#
#####################################################################################################
#####################################################################################################
#									update channel information										#
#####################################################################################################
def updateChannel():
	fileName = inputExcelFileName()
	while len(fileName)==0:
		fileName = inputExcelFileName()
	startLine = inputImportStartLine()
	while startLine < 0:
		startLine = inputImportStartLine()
	srcFile = "tmp/" + fileName

	simpleFields=["channel_type","channel_classification","channel_manager","channel_manager_phone","manage_college_count",
	"recruitstudents_principal","recruitstudents_principal_phone","recruitstudents_principal_tel","recruitstudents_principal_email",
	"senate_principal","senate_principal_phone","senate_principal_tel","senate_principal_email","main_items","cooperation_items"]
	complexFileds=["channel_name","corporation_name","docking_people_id","channel_phone","channel_postcode","channel_officeaddress"]
	content = getContentFromXls(srcFile)
	worksheet = content.sheets()[0]
	total = worksheet.nrows - startLine
	print("analyze the file... total records sums to " + str(total))
	for i in range(startLine,worksheet.nrows):
		progressbar(i - startLine,total)
		row = worksheet.row_values(i)
		for j in xrange(0,worksheet.ncols):
			value = row[j]
			ctype = worksheet.cell(i,j).ctype
			if ctype==0:
				pass
			elif ctype==1:
				value = value.strip()
			elif ctype==2:
				pass
			elif ctype==3:
				value = xlrd.xldate.xldate_as_datetime(value,1)
			elif ctype==4:
				pass
			else:
				pass
			if j==0:
				field = value
			elif j==1:
				val = value
			elif j==2:
				name = value
			else:
				pass
		count1 = simpleFields.count(field)
		count2 = complexFileds.count(field)
		#print (field)
		#print ("count1:" + str(count1) + "/count2:" + str(count2))
		sqls = []
		if count1 > 0:
			sql = "update channel set " + field + "="
			if field=="manage_college_count":
				sql = sql + str(val)
			else:
				sql = sql + "'" + val + "'"
			sql = sql + " where channel_name='" + name + "'"
			#print(sql)
			sqls.append(sql)
		elif count2 > 0:
			if field=="channel_name":
				sql = "select channel_id,channel_userid from channel where channel_name='" + name + "'"
				recs = selectFromDB(glovar.db4xf,sql)
				if len(recs) == 0:
					break
				cid = recs[0][0]
				cuid = recs[0][1]
				print("cid=" + str(cid) + "/cuid=" + str(cuid))
				sql = "update channel_protocol set channel_name='" + val + "' where channel_id=" + str(cid)
				sqls.append(sql)
				sql = "update \"user\" set user_name='" + val + "' where user_id=" + str(cuid)
				sqls.append(sql)
				sql = "update channel set channel_name='" + val + "' where channel_id=" + str(cid)
				sqls.append(sql)
				sql = "update student_enroll set student_source='" + val + "' where student_source='" + name + "'"
				sqls.append(sql)
			elif field=="docking_people_id":
				sql = "select user_id from \"user\" where user_name='" + val + "' and account_status='1'"
				recs = selectFromDB(glovar.db4xf,sql)
				if len(recs)==0:
					pass
				else:
					uid = recs[0][0]
					sql = "update channel set user_id=" + str(uid) + ",docking_people_id=" + str(uid) + " where channel_name='" + name + "'"
					sqls.append(sql)
					sql = "update student_enroll set enrollment_people='" + val + "' where student_id in (select s.student_id from "
					sql = sql + "student AS s LEFT JOIN student_enroll AS se ON s.student_id = se.student_id where s.enroll_batch='" + glovar.currentBatch + "' "
					sql = sql + "and se.student_source='" + name + "')"
					print(sql)
					sqls.append(sql)
			else:
				sql = "update channel_protocol set " + field + "='" + val + "' where channel_name='" + name + "'"
				sqls.append(sql)
				sql = "update channel set " + field + "='" + val + "' where channel_name='" + name + "'"
				sqls.append(sql)				
				pass
			#print (sql)
		else:
			pass
		#print(len(sqls))
		if len(sqls)>0:
			executeFromDB(glovar.db4xf,sqls)
	print("")
	mainMenu()
#####################################################################################################
#									update channel information										#
#####################################################################################################
#select function entrance
def selectFunc():
    func = -1;
    try:
	    func = int(raw_input("select function(1-CheckImport;2-Statistics,3-Update Attachment,4-Update student,5-Update channel): "))
    except Exception, e:
    	func = -1
    else:
    	if func<1 or func>5:
    		func = -1
    finally:
    	pass
    return func

def mainMenu():
	func = selectFunc()
	while func < 0:
		func = selectFunc()
	if func == 1:
		checkImport()
	elif func == 2:
		getEnrollDataForStatistics()	
	elif func==3:
		updateChannelAttach()
	elif func==4:
		updateStudent()
	elif func==5:
		updateChannel()
	pass
#####################################################################################################
#											Program Entrance										#
#####################################################################################################
initialize()
mainMenu()
#####################################################################################################
#											Program Entrance										#
#####################################################################################################