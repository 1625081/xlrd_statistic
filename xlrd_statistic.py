import xlrd
import os
import csv
import json
class LastCharacterError(Exception):
    def __init__(self):
        Exception.__init__(self)
#请自行修改以下的path变量
path="C:\\Users\\ZhengJinYang\\Downloads\\Classrooms\\"
#复制文件所在地址，并使用\\代替\
try:
    dirs=os.listdir(path)
    if path[-1]!="\\" and path[-1]!="/":
        raise LastCharacterError
except FileNotFoundError:
    print("运行错误：请检查文件目录(path变量)是否正确！")
    exit(1)
except LastCharacterError:
    print("Windows系统下,目录必须以\\结尾;Unix下必须以/结尾")
    exit(1)
counter=[[[] for j in range(5)] for i in range(6)]
classes={}

for xls in dirs:
    try:
        data=xlrd.open_workbook(path+xls)
    except:
        print("运行错误：")
        print("1.请检查目录中文件是否都为xls格式")
        print("2.没有文件已在excel中打开")
        exit(1)
    table=data.sheets()[0]
    for i in range(table.nrows):
        row=table.row_values(i)
        if i==0:
            a=row[0].split(")")[1]
            stu=a.split("课")[0]
        elif i>1 and i<8:
            for j in range(2,7):
                if row[j]!='':
                    counter[i-2][j-2].append(stu)
                    if str(i)+" "+str(j)+":"+row[j] in classes:
                        classes[str(i)+" "+str(j)+":"+row[j]].append(stu)
                    else:
                        classes[str(i)+" "+str(j)+":"+row[j]]=[stu]
json_result=json.dumps(classes,sort_keys=True,indent=4,separators=(',',':'),ensure_ascii=False)
f=open(path+'json_result.txt','a+')
f.write(json_result)

with open(path+'_result.csv','a+') as csvfile:
    writer = csv.writer(csvfile,dialect='excel')
    writer.writerow(["周一","周二","周三","周四","周五"])
    for i in range(6):
        writer.writerow(counter[i])
with open(path+'_numresult.csv','a+') as csvfile:
    writer = csv.writer(csvfile,dialect='excel')
    writer.writerow(["周一","周二","周三","周四","周五"])
    for i in range(6):
        for j in range(5):
            counter[i][j]=len(counter[i][j])
    for i in range(6):
        writer.writerow(counter[i])
