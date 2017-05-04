#!/usr/bin/python
#-*-coding:UTF-8-*-
import psycopg2,sys,math

#select function
def selectFromDB(dbParam,sql):
    #print(dbParam["database"] + "/" + dbParam["user"] + "/" + dbParam["password"] + "/" + dbParam["host"])
    conn = psycopg2.connect(database=dbParam["database"],user=dbParam["user"],password=dbParam["password"],host=dbParam["host"],port=dbParam["port"])
    cur = conn.cursor()
    cur.execute(sql)
    rows = cur.fetchall()
    conn.close()  
    return rows

#transform suspiciousRecord into excel
def transformSusRecord2Xls():
    xlsStr = ""
    return xlsStr

#transform suspiciousRecord into csv
def transformSusRecord2Csv(records):
    csvStr = ""
    if len(records) > 0:
        csvStr = "identity_card,student_name,study_center_name,enroll_school,major,enroll_arrangement\n"
        
        for ids,rec1 in records.items():
            csvStr = csvStr + "'" + ids + "',"
            csvStr = csvStr + rec1.get("student_name","") + ","
            csvStr = csvStr + rec1.get("study_center_name","") + ","
            csvStr = csvStr + rec1.get("enroll_school","") + ","
            csvStr = csvStr + rec1.get("major","") + ","
            csvStr = csvStr + rec1.get("enroll_arrangement","") + "\n"
            print(csvStr)
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

def progressbar(cur, total):
    percent = '{:.2f}'.format(float(cur) / float(total))
    sys.stdout.write('\r')
    sys.stdout.write("[%-50s] %s" % ('=' * int(math.floor(cur * 50 / total)),percent))
    sys.stdout.flush()

#output type
outputType = selectOutputType();
while outputType < 0:
    outputType = selectOutputType()

#database = "testdb", user = "uuer", password = "ppas", host = "127.0.0.1"
db4xf = {"database":"education","user":"myedu","password":"dbpassword","host":"172.16.40.4","port":"3433"}
reduplicative = {}
nonIds = 0

#load data from xfu
sql = "select s.student_name,s.identity_card,se.study_center_name,se.enroll_school,se.major,se.enroll_arrangement from student as s "
sql = sql + "left join student_enroll as se on s.student_id=se.student_id where 1=1 and s.is_delete_flag is null "
sql = sql + "and s.enroll_batch='1703' order by s.identity_card desc"
#print(sql)
print("get data from xuefu database...")
rows = selectFromDB(db4xf,sql)
total = len(rows)
i = 0
print("finding duplicate identity_card records in dta... data size is:" + str(total))
for i in xrange(0,len(rows) - 1):
    progressbar(i,total)
    row1 = rows[i]
    row2 = rows[i + 1]
    if row1[1] == None:
        ids = "--" + str(nonIds)
        nonIds = nonIds + 1;
    else:
        ids = row1[1]
    idCard1 = ids
    if row2[1] == None:
        ids = "--" + str(nonIds)
        nonIds = nonIds + 1;
    else:
        ids = row2[1]
    idCard2 = ids
    #print(idCard1)
    rec = {}
    if idCard1 == idCard2:
        if row1[0] == None:
            rec["student_name"] = ""
        else:
            rec["student_name"] = row1[0]
        if row1[2] == None:
            rec["study_center_name"] = ""
        else:
            rec["study_center_name"] = row1[2]
        if row1[3] == None:
            rec["enroll_school"] = ""
        else:
            rec["enroll_school"] = row1[3]
        if row1[4] == None:
            rec["major"] = ""
        else:
            rec["major"] = row1[4]
        if row1[5] == None:
            rec["enroll_arrangement"] = ""
        else:
            rec["enroll_arrangement"] = row1[5]
        reduplicative[idCard1] = rec
        tmpstr = idCard1 + "/" + rec["student_name"]+"/"+rec["study_center_name"]+"/"+rec["enroll_school"]+"/"+rec["major"]+"/"+rec["enroll_arrangement"]
        #print(tmpstr)
print("")
outputStr = ""
fileName = "duplicate_identity_card"

if outputType == 1:
    outputStr = transformSusRecord2Xls(reduplicative)
    fileName = fileName + ".xls"
elif outputType == 2:
    outputStr = transformSusRecord2Csv(reduplicative)
    fileName = fileName + ".csv"
elif outputType == 3:
    outputStr = transformSusRecord2Json(reduplicative)
    fileName = fileName + ".json"
else:
    pass
file_path = "result/" + fileName
fw = open(file_path,"w")
fw.write(outputStr)
fw.close()    
