#!/usr/bin/python
#-*-coding:UTF-8-*-
import os,psycopg2,sys,stat,math
import mssql

#select function
def selectFromDB(dbParam,sql):
    #print(dbParam["database"] + "/" + dbParam["user"] + "/" + dbParam["password"] + "/" + dbParam["host"])
    conn = psycopg2.connect(database=dbParam["database"],user=dbParam["user"],password=dbParam["password"],host=dbParam["host"],port=dbParam["port"])
    cur = conn.cursor()
    cur.execute(sql)
    rows = cur.fetchall()
    conn.close()  
    return rows

#Dictionary compare, equal returns 0, nonEqual returns -1
def compareRec(rec1,rec2):
    ret = 0;
    for rKey,rValue in rec1.items():
        r2Value = rec2[rKey]
        if rValue != r2Value:
            ret = -1
            break
        else:
            pass
    return ret
#transform suspiciousRecord into excel
def transformSusRecord2Xls():
    xlsStr = ""
    return xlsStr

#transform suspiciousRecord into csv
def transformSusRecord2Csv(records):
    csvStr = ""
    if len(records) > 0:
        csvStr = "identity_card,student_name,study_center_name,enroll_school,major,enroll_arrangement"
        csvStr = csvStr + ",student_name,study_center_name,enroll_school,major,enroll_arrangement\n"
        total = len(records)
        print("starting transform array... data size is:" + str(total))
        i = 0
        for ids,susRecList in records.items():
            progressbar(i,total)
            i = i + 1
            rec1 = susRecList[0]
            rec2 = susRecList[1]
            csvStr = csvStr + "'" + ids + "',"
            csvStr = csvStr + rec1.get("student_name","") + ","
            csvStr = csvStr + rec1.get("study_center_name","") + ","
            csvStr = csvStr + rec1.get("enroll_school","") + ","
            csvStr = csvStr + rec1.get("major","") + ","
            csvStr = csvStr + rec1.get("enroll_arrangement","") + ","
            csvStr = csvStr + rec2.get("student_name","") + ","
            csvStr = csvStr + rec2.get("study_center_name","") + ","
            csvStr = csvStr + rec2.get("enroll_school","") + ","
            csvStr = csvStr + rec2.get("major","") + ","
            csvStr = csvStr + rec2.get("enroll_arrangement","") + "\n"
            #print(csvStr)
        print("")
    else:
        pass
    return csvStr

#transform suspiciousRecord into json
def transformSusRecord2Json():
    jsonStr = ""
    return jsonStr

#select output type
def selectOutputType():
    sel = -1;
    inp = int(raw_input("Select output format(1-excel;2-cvs;3-json):  "))
    if inp < 1:
        pass
    elif inp >3:
        pass
    else:
        sel = inp
    return sel

#select Data scope
def selectDataScope():
    sel = -1;
    inp = int(raw_input("Select data scope(1-All;2-StudyCenter;3-Univercity;4-branch):  "))
    if inp < 1:
        pass
    elif inp >4:
        pass
    else:
        sel = inp
    return sel

def progressbar(cur, total):
    percent = '{:.2f}'.format(float(cur) / float(total))
    sys.stdout.write('\r')
    sys.stdout.write("[%-50s] %s" % ('=' * int(math.floor(cur * 50 / total)),percent))
    sys.stdout.flush()

def findInSortList(val,tList,direction):
    #direction indicates tList sorting, 1 for asc,0 for desc
    found = -1
    for i in range(len(tList)):
        value = tList[i]
        if val==value:
            found = i
            break
        elif val<value:
            if direction==1:
                continue
            else:
                break
        else:
            if direction==0:
                break
            else:
                continue
    return found

def compare2SortedList(list1,list2,direction):
    list2bi = 0;
    i = 0;
    k = -1;
    total1 = len(list1)
    for i in range(total1):
        progressbar(k,len(list1))
        k = k + 1
        val1 = list1[k]
        for j in xrange(list2bi,len(list2)):
            val2 = list2[j]
            if val1==val2:
                list1.pop(k)
                k = k - 1
                list2.pop(j)
                listbi=j
                break
            elif val1<val2:
                if direction==1:
                    break
            else:
                if direction==0:
                    break
    print("")

#database = "testdb", user = "uuer", password = "ppas", host = "127.0.0.1"
db4ds = {"database":"zhijin_data","user":"myedu","password":"dbpassword","host":"172.16.40.4","port":"3433"}
db4xf = {"database":"education","user":"myedu","password":"dbpassword","host":"172.16.40.4","port":"3433"}
db4es = {"database":"eszjedu","user":"cnxfu","password":"My-es-password","host":"172.16.30.19","port":"3433"}
suspiciousRecord = {}
xfLostRecord = {}
dsRecords = {}
xfRecords = {}
esRecords = []
esRecords1 = []
esLearningCenter = []
xf2Records = []
xf2Records1 = []
xfStudyCenter = []
esDiffxf = []
nonIds = 0

reload(sys)  
sys.setdefaultencoding('utf8')   

#output type
outputType = selectOutputType();
while outputType < 0:
    outputType = selectOutputType()
#select data scope to check
key = ''
dataScope = selectDataScope()
while dataScope < 0:
    dataScope = selectDataScope()
dsValue = None
if dataScope == 2:
    dsValue = "'" + raw_input("Please input study center name:  ") + "'"
elif dataScope == 3:
    dsValue = "'" + raw_input("please input univercity name:  ") + "'"
elif dataScope == 4:
    dsValue = "'" + raw_input("please input branch name:   ") + "'"
else:
    pass

#tstr = "outputType:" + str(outputType) + ",dataScope:" + str(dataScope)
#if dataScope == 2:
#    tstr = tstr + "study_center_name:" + dsValue
#elif dataScope == 3:
#    tstr = tstr + "univercity:" + dsValue
#else:
#    pass
#print(tstr)

#load data from ES

sql = "select certificateno,LearningCenterName from tb_studentbaseinfo where recruitbatchname='1703' "
if dsValue!=None:
    sql = sql + "and HeadStationName=" + dsValue + " "
sql = sql +  "order by certificateno"
print(sql)
ms = mssql.MSSQL(host=db4es["host"],port=db4es["port"],user=db4es["user"],pwd=db4es["password"],db=db4es["database"])
resList = ms.ExecQuery(sql)
print("loading data from ES... data size is:" + str(len(resList)))
for i in range(len(resList)):
    progressbar(i,len(resList))
    rec = resList[i]
    if rec[0]=='':
        print (rec[0])
    esRecords.append(rec[0])
    esRecords1.append(rec[0])
    esLearningCenter.append(rec[1])
print("")
#load data from xf audit_status=2
sql = "select s.identity_card,se.study_center_name from student as s left join student_enroll as se on s.student_id=se.student_id "
sql = sql + "left join \"user\" as u on s.user_id=u.user_id "
sql = sql + "where se.audit_status='2' and s.is_delete_flag is null "
sql = sql + "and s.user_id is not null and s.enroll_batch='1703' "
if dsValue!=None:
    sql = sql + "and u.standby2=" + dsValue + " "
sql = sql + "order by identity_card"
print (sql)
rows = selectFromDB(db4xf,sql)
total = len(rows)
print("loading data from xf, with audit_status=2... data size is:" + str(total))
for i in range(total):
    progressbar(i,total)
    rec = rows[i]
    xf2Records.append(rec[0])
    xf2Records1.append(rec[0])
    xfStudyCenter.append(rec[1])
print("")

#load data from zhijin_data
sql = "select student_name,identity_card,study_center_name,enroll_school,major,enroll_arrangement from crawler_student "
sql = sql + "where 1=1 and status=1 and enroll_batch='1703' "
if dataScope == 2:
    sql = sql + "and study_center_name=" + dsValue
elif dataScope == 3:
    sql = sql + "and enroll_school=" + dsValue
else:
    pass
sql = sql + " order by identity_card desc"
#print (sql)
rows = selectFromDB(db4ds,sql)
total = len(rows)
print("loading data from data source... data size is:" + str(total))
i = 0
for row in rows:
    #print(row[0]+"/"+row[1]+"/"+row[2]+"/"+row[3]+"/"+row[4]+"/"+row[5])
    progressbar(i,total)
    i = i + 1
    rec = {}
    ids = ""
    if row[0] == None:
        rec["student_name"] = ""
    else:
        rec["student_name"] = row[0]
    if row[2] == None:
        rec["study_center_name"] = ""
    else:
        rec["study_center_name"] = row[2]
    if row[3] == None:
        rec["enroll_school"] = ""
    else:
        rec["enroll_school"] = row[3]
    if row[4] == None:
        rec["major"] = ""
    else:
        rec["major"] = row[4]
    if row[5] == None:
        rec["enroll_arrangement"] = ""
    else:
        rec["enroll_arrangement"] = row[5]
    #print(rec["student_name"]+"/"+rec["study_center_name"]+"/"+rec["enroll_school"]+"/"+rec["major"]+"/"+rec["enroll_arrangement"])
    if row[1] == None:
        ids = "--" + str(nonIds)
        nonIds = nonIds + 1;
    else:
        ids = row[1]
    dsRecords[ids] = rec
print("")

#load data from xfu
sql = "select s.student_name,s.identity_card,se.study_center_name,se.enroll_school,se.major,se.enroll_arrangement from student as s "
sql = sql + "left join student_enroll as se on s.student_id=se.student_id where 1=1 and s.is_delete_flag is null "
sql = sql + "and s.enroll_batch='1703' "
if dataScope == 2:
    sql = sql + "and se.study_center_name" + dsValuoe
elif dataScope == 3:
    sql = sql + "and se.enroll_school=" + dsValue
else:
    pass
sql = sql + " order by s.identity_card desc"
#print (sql)
rows = selectFromDB(db4xf,sql)
total = len(rows)
print("loading data from xuefu... data size is:" + str(total))
i = 0
for row in rows:
    #print(row[0]+"/"+row[1]+"/"+row[2]+"/"+row[3]+"/"+row[4]+"/"+row[5])
    progressbar(i,total)
    i = i + 1
    rec = {}
    ids = ""
    if row[0] == None:
        rec["student_name"] = ""
    else:
        rec["student_name"] = row[0]
    if row[2] == None:
        rec["study_center_name"] = ""
    else:
        rec["study_center_name"] = row[2]
    if row[3] == None:
        rec["enroll_school"] = ""
    else:
        rec["enroll_school"] = row[3]
    if row[4] == None:
        rec["major"] = ""
    else:
        rec["major"] = row[4]
    if row[5] == None:
        rec["enroll_arrangement"] = ""
    else:
        rec["enroll_arrangement"] = row[5]
    #print(rec["student_name"]+"/"+rec["study_center_name"]+"/"+rec["enroll_school"]+"/"+rec["major"]+"/"+rec["enroll_arrangement"])
    if row[1] == None:
        ids = "--" + str(nonIds)
        nonIds = nonIds + 1;
    else:
        ids = row[1]
    xfRecords[ids] = rec
print("")
#compare data source with xuefu
#if xfRecords doesn't have the record in dsRecords or the record in xfRecords doesn't match the dsRecords(they have the same identity_card)
total = len(dsRecords)
print("comparing dsRecords and xfRecords... data size is:" + str(total))
i = 0
for id1,rec1 in dsRecords.items():
    progressbar(i,total)
    i = i + 1
    rec2 = xfRecords.get(id1,{"student_name":"","study_center_name":"","enroll_school":"","major":"","enroll_arrangement":""})
    compres = compareRec(rec1,rec2)
    if compres == -1:
        recNull = {"student_name":"","study_center_name":"","enroll_school":"","major":"","enroll_arrangement":""}
        compres = compareRec(rec2,recNull)
        if compres == -1:
            suspiciousRecord[id1] = [rec1,rec2]
        else:
            xfLostRecord[id1] = [rec1,rec2]
print("")

#compare xfu and es
print("finding short in ES and XF...")
compare2SortedList(esRecords,xf2Records,1)
xfBegin = 0
total = len(esRecords1)
print("finding diffrent between es and xf... data size is " + str(total))
for i in range(total):
    progressbar(i,total)
    esID = esRecords1[i]
    esLC = esLearningCenter[i]
    for j in xrange(xfBegin,len(xf2Records1)):
        xfID = xf2Records1[j]
        xfSC = xfStudyCenter[j]
        print(str(i) + ":" + esID + " | " + str(j) + ":" + xfID)
        if esID==xfID:
            print (esID)
            if esLC==xfSC:
                print("\n")
                pass
            else:
                print (esID + "\tes:" + esLC + "\txf:" + xfSC + "\n")
                esDiffxf.append(esID + "\tes:" + esLC + "\txf:" + xfSC)
            xfBegin = j
            break;
        elif esID>xfID:
            continue
        else:
            xfBegin = j
            break
        x = raw_input("wait")
print ("")

print("output result...")
outputStr1 = ""
outputStr2 = ""
fileName1 = "diff_between_ds_xf"
fileName2 = "ds_has_xf_nil"
if dataScope == 2 or dataScope == 3:
    fileName1 = fileName1 + "-" + dsValue[1:len(dsValue) - 1]
    fileName2 = fileName2 + "-" + dsValue[1:len(dsValue) - 1]
else:
    pass
if outputType == 1:
    outputStr1 = transformSusRecord2Xls(suspiciousRecord)
    outputStr2 = transformSusRecord2Xls(xfLostRecord)
    fileName1 = fileName1 + ".xls"
    fileName2 = fileName2 + ".xls"
elif outputType == 2:
    outputStr1 = transformSusRecord2Csv(suspiciousRecord)
    outputStr2 = transformSusRecord2Csv(xfLostRecord)
    fileName1 = fileName1 + ".csv"
    fileName2 = fileName2 + ".csv"
elif outputType == 3:
    outputStr1 = transformSusRecord2Json(suspiciousRecord)
    outputStr2 = transformSusRecord2Json(xfLostRecord)
    fileName1 = fileName1 + ".json"
    fileName2 = fileName2 + ".json"
else:
    pass
file_path = "result/" + fileName1
fw = open(file_path,"w")
fw.write(outputStr1)
fw.close()
file_path = "result/" + fileName2
fw = open(file_path,"w")
fw.write(outputStr2)
fw.close()
outputStr2 = ""
#for i in range(len(shortInES)):
for i in range(len(esRecords)):
    id1 = esRecords[i]
    outputStr2 = outputStr2 + "'" + id1 + "',\n"
outputStr2 = outputStr2.strip()
outputStr2 = outputStr2[:len(outputStr2) - 1]
file_path = "result/shortInXF"
fw = open(file_path,"w")
fw.write(outputStr2)
fw.close()
outputStr2 = ""
#for i in range(len(shortInXF)):
for i in range(len(xf2Records)):
    id1 = xf2Records[i]
    outputStr2 = outputStr2 + "'" + id1 + "',\n"
outputStr2 = outputStr2.strip()
outputStr2 = outputStr2[:len(outputStr2) - 1]
file_path = "result/shortInES"
fw = open(file_path,"w")
fw.write(outputStr2)
fw.close()

outputStr2 = ""
for i in range(len(esDiffxf)):
    id1 = esDiffxf[i]
    outputStr2 = outputStr2 + "'" + id1 + "'\n"
outputStr2 = outputStr2.strip()
file_path = "result/esDiffxf"
fw = open(file_path,"w")
fw.write(outputStr2)
fw.close()