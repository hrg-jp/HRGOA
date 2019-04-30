#!/usr/bin/python
# -*- coding: UTF-8 -*-
import sys
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
import MyEmail
from Json2Dict import json2dict
from os.path import dirname, abspath,join
import pandas as pd
import copy
private_path=join(dirname(dirname(abspath(__file__))),"private")
print(private_path)
service = MyEmail.QQEXMailService(join(private_path,r"email-account.json"))

template = MyEmail.EmailTemplate(join(private_path,r"email-template.txt"),
                                 join(private_path,r"email-content.case01.txt"))
def readTasks():
    try:
        dfTask = pd.read_excel(join(private_path,r"Tasks.xlsx"),sheet_name="CurrentTask",index_col=[0],converters={'Result':str})
        return dfTask
    except:
        return None
def getPeddingTaskCount(dfTask):
    try:
        return len(dfTask[dfTask["Status"]=="Pedding"])
    except:
        return 0

def init_task():
    try:
        df_raw = pd.read_excel(join(private_path,r"Customers.xlsx"),sheet_name="DataSource",index_col=[0])
        df_raw["Status"]= "Pedding"
        df_raw["Result"]= ""
        df_raw["FinishTime"]= ""
        df_raw["CustomerEmail"]= df_raw["メール"]
        df_raw["CustomerCompanyName"]= df_raw["会社名"]
        df_raw["CustomerNameinHeader"]= df_raw["担当"]
        df_raw["CustomerNameinContent"]= df_raw["担当"]
        del df_raw["担当"]
        del df_raw["会社名"]
        del df_raw["電話番号"]
        del df_raw["微信"]
        del df_raw["メール"]
        del df_raw["ホームページ"]
        del df_raw["住所"]
        df_raw.to_excel(join(private_path,r"Tasks.xlsx"),sheet_name="CurrentTask")
    except:
        print("打开数据时发生错误，程序异常关闭")
        sys.exit()
def sending(dfTask):    
    default_params = json2dict(join(private_path,r"email-params.json"))
    injection_cols=["CustomerEmail","CustomerCompanyName","CustomerNameinHeader","CustomerNameinContent"]
    print("Starting sending...total:{}".format(getPeddingTaskCount(dfTask)))
    for i,row in dfTask.iterrows():
        if row["Status"]!="Pedding":
            continue
        params = copy.deepcopy(default_params)
        
        for injection_col in injection_cols:
            value = str(row[injection_col])
            if len(value.strip())>0 and value!="nan":
                params[injection_col] = row[injection_col]

        email = MyEmail.EmailEntity(template,params)

        retry_time = 3
        result = False
        while(retry_time>0):
            result = service.sendEmail(email)
            if(result):
                break
        dfTask.at[i,"Result"] = "{}".format(result)
        dfTask.at[i,"Status"] = "Finished"
        dfTask.at[i,"FinishTime"] = pd.Timestamp.now()
        print("[{Result}]Sending to {CustomerEmail},{CustomerCompanyName},{CustomerNameinHeader}".format(
            Result=result,**params))
        dfTask.to_excel(join(private_path,r"Tasks.xlsx"),sheet_name="CurrentTask")
dfTask = readTasks()
PeddingTaskCount = getPeddingTaskCount(dfTask)

if(PeddingTaskCount>0):
    choise = input("您有未完成的任务{}个，是否继续发送？Y/N".format(PeddingTaskCount))
    if(choise.strip().upper()=="Y"):
        sending(dfTask)
        print("任务完成")
        sys.exit()
    elif(choise.strip().upper()=="N"):
        pass
    else:
        print("退出中。。")
        sys.exit()
else:  
    choise = input("当前没有邮件任务，是否开始新的邮件任务？Y/N")
    if(choise.strip().upper()=="Y"):
        init_task()
        dfTask = readTasks()
        sending(dfTask)
        print("任务完成")
        sys.exit()
    else:
        print("退出中。。")
        sys.exit()