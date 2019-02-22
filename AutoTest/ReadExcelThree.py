import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.text import MIMEText
import datetime
import requests
import xlrd
import xlwt
import xlutils3
from xlutils3 import copy

#文件位置
ExcelFile=xlrd.open_workbook(r'D:\Program Files (x86)\测试用例.xlsx')
#获取目标EXCEL文件sheet名
# print(ExcelFile.sheet_names()[0])
sheet=ExcelFile.sheet_by_name('Sheet1')
nrows=sheet.nrows-1

#复制EXCEL文件
wtbook = copy.copy(ExcelFile)
wtsheet = wtbook.get_sheet(0)
type(wtsheet)

nowTime=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')#现在

for i in range(nrows):
   cols=sheet.col_values(2).__getitem__(i+1) #获取request
   response=sheet.col_values(3).__getitem__(i+1)

   cookies = dict(IsLogin='0',ZDCustomerToken='4eafb7da-6e9f-49bd-b508-1f480be4160f',CustomerSysNo='12879',userGuid='5bd897ab-fce8-4166-90ef-742d0a4b5bc820180828014901')
   #url="httpsheet://192.168.60.228:40011/IAccount/Login?customerID=15221842558&LoginType=2&md5Pwd=e10adc3949ba59abbe56e057f20f883e&Version=3.0.1&SystemVersion=10.3.1&PhoneModel=iPhone&Source=1&DeviceID=00244EE1-63CA-4FB1-826A-A907D27196D2&sign=28e532b69db923583fc562508af45bf2"
   actualresponse = requests.request("GET", cols,cookies=cookies)
   actualresponse1= actualresponse.text
   if response[10]==actualresponse1[10]:
        wtsheet.write(i + 1, 4, actualresponse1[10])
        wtsheet.write(i + 1, 5, actualresponse1[24]+actualresponse1[25]+actualresponse1[26]+actualresponse1[27])
        wtsheet.write(i+1, 6, 'sucess')
        wtsheet.write(i + 1, 7, nowTime)
   else:
       wtsheet.write(i + 1, 4, actualresponse1[10])
       wtsheet.write(i + 1, 5, actualresponse1[24]+actualresponse1[25]+actualresponse1[26]+actualresponse1[27])
       wtsheet.write(i + 1, 6, 'fail')
       wtsheet.write(i + 1, 7, nowTime)
wtbook.save('D:\Program Files (x86)\TestCase\TestCaseResult.xls')

# 指定测试报告的路径
report_dir = 'D:\\Program Files (x86)\\TestCase'

def new_file(test_dir):
    #列举test_dir目录下的所有文件，结果以列表形式返回。
    lists=os.listdir(test_dir)
    #sort按key的关键字进行排序，lambda的入参fn为lists列表的元素，获取文件的最后修改时间
    #最后对lists元素，按文件修改时间大小从小到大排序。
    #lists.sort(key=lambda fn:os.path.getmtime(test_dir+'\\'+fn))
    #获取最新文件的绝对路径
    file_path=os.path.join(test_dir,lists[-1])
    return file_path


# 发送邮件，发送最新测试报告html
def send_email(newfile):
    # 打开文件
    f = open(newfile, 'rb')
    # 读取文件内容
    mail_body = f.read()
    # 关闭文件
    f.close()

    # 发送邮箱服务器
    smtpserver = 'smtp.qq.com'
    # 发送邮箱用户名/密码
    user = '1378915244@qq.com'
    password = 'wegydhtjryijhcgd'
    # 发送邮箱
    sender = '1378915244@qq.com'
    # 多个接收邮箱，单个收件人的话，直接是receiver='XXX@163.com'
    receiver = ['1378915244@qq.com']

    #发送主题及附件
    msg = MIMEMultipart()
    msg['Subject'] = Header('Python接口自动化测试报告', 'utf-8')
    msg_excel1 = MIMEText(mail_body, 'excel', 'utf-8')
    msg.attach(msg_excel1)
    msg_excel = MIMEText(mail_body, 'excel', 'utf-8')
    msg_excel["Content-Disposition"] = 'attachment; filename="TestCaseResult.xls"'
    msg.attach(msg_excel)

    # 要加上msg['From']这句话，否则会报554的错误。
    # 要在163设置授权码（即客户端的密码），否则会报535
    msg['From'] = '1378915244@qq.com'
    # 多个收件人
    msg['To'] = "1378915244@qq.com"

    # 连接发送邮件
    smtp = smtplib.SMTP()
    smtp.connect(smtpserver)
    smtp.login(user, password)
    smtp.sendmail(sender, receiver, msg.as_string())
    smtp.quit()

new_report = new_file(report_dir)
send_email(new_report)