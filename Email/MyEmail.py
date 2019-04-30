import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
import re
import json
class EmailEntity(object):
    def __init__(self,emailTemplate,params, *args, **kwargs):
        self.params = params
        self.title = emailTemplate.title.format(**params)
        self.content = emailTemplate.content.format(**params)
        return super().__init__(*args, **kwargs)
class EmailTemplate(object):
    def __init__(self,template_file,content_case_file, *args, **kwargs):
        with open(template_file,encoding="utf-8") as file:
            raw = file.readline();
            self.title = re.search(r'Title:(.*)',string=raw,flags=re.IGNORECASE).group(1)
            raw = file.readlines()
            self.content = re.search(r'Content:([\s\S]*)',string="".join(raw),flags=re.IGNORECASE).group(1)
        raw = open(content_case_file,encoding="utf-8").readlines()
        self.content = self.content.format(CaseContent="".join(raw))
        return super().__init__(*args, **kwargs)
class QQEXMailService(object):
    def __init__(self,json_file, *args, **kwargs):
        with open(json_file) as json_data:
            emailconfig = json.load(json_data)
            json_data.close()
        self.Account = emailconfig['Account']
        self.Password = emailconfig['Password']
        #self.SenderEmail = emailconfig['SenderEmail']
        #self.SenderNameinHeader = emailconfig['SenderNameinHeader']
        #self.SenderNameinContent = emailconfig['SenderNameinContent']
        #self.SenderNameinContentNoSpace = emailconfig['SenderNameinContentNoSpace']
        return super().__init__(*args, **kwargs)
    def sendEmail(self,emailEntity):
        ret=True
        try:
            params = emailEntity.params
            msg=MIMEText(emailEntity.content,'plain','utf-8')
            msg['From']=formataddr([params["SenderNameinHeader"],params["SenderEmail"]])
            msg['To']=formataddr([params["CustomerNameinHeader"],params["CustomerEmail"]])
            msg['Subject']=emailEntity.title
 
            server=smtplib.SMTP_SSL("smtp.exmail.qq.com", 465)  # 发件人邮箱中的SMTP服务器，端口是25
            server.login(self.Account, self.Password)  # 括号中对应的是发件人邮箱账号、邮箱密码
            server.sendmail(params["SenderEmail"],[params["CustomerEmail"],],msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
            server.quit()  # 关闭连接
        except Exception:  # 如果 try 中的语句没有执行，则会执行下面的 ret=False
            ret=False
        return ret


