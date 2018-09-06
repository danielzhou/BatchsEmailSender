import os
import sys
import csv
import smtplib
from email.mime.base import MIMEBase 
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email import encoders

# ========================批量发送邮件测试（二）----邮件内容固定，主题和附件变化=================================


# --------------------发送服务器配置---------------
sender_host = 'smtp.163.com:25'  # 默认服务器地址及端口
sender_user = 'laoliu1810181@163.com'  
sender_pwd = '123456zwj'
sender_name = u'上海周YJ公司'

attach_path = r'C:\Users\user\Desktop\MailMaster V1.1\attchfile'   # 附件所在文件夹
attach_type = ".xlsx"      # 附件后缀名，即类型
addrBook = r'C:\Users\user\Desktop\MailMaster V1.1\邮箱联系人表单.csv'  # 邮件地址通讯录
content_path = r"C:\Users\user\Desktop\MailMaster V1.1\content.txt"   # 邮箱正文内容.txt

# --------------根据输入的CSV文件，获取通讯录人名和相应的邮箱地址-------

def getAddrBook(addrBook):
	'''
		@作用：根据输入的CSV文件，形成相应的通讯录字典
		@返回：字典类型，name为人名，value为对应的邮件地址
	'''
	with open(addrBook,'r',encoding='gbk') as addrFile: 
		reader = csv.reader(addrFile)
		name = []
		value = []
		for row in reader:
			name.append(row[0])
			value.append(row[1])
	
	addrs = dict(zip(name, value))
	return addrs

# addrs = {name : value}


# ----------获取所有附件 及其不带后缀的文件名（作为主题）---------

# 例如，输入：C:\Users\zwj\Desktop\MailMaster\attchfile\结算款_1_周.xlsx
# 输出：结算款_1_周
# def getAttNameList(attach_path,filetype):
# 	'''
# 		@作用：输入文件夹 和文件类型，获得该文件夹下的所有该类型文件
# 		@返回：1. 该类型文件的文件名列表，不包括后缀和路径
# 			  2. 该类型文件转化为邮箱附件类型后的列表
# 	'''
# 	att_name = []
# 	att = []
# 	for root,dirs,files in os.walk(attach_path):
# 		for i in files:
# 			if filetype in i:
# 				att_name.append(i.replace(filetype,""))  # 生成不带‘.xlsx’后缀名的文件名
# 	return att_name

# att_name = getAttNameList(attach_path,attach_type)

# ---------------根据对应的附件名，获取要发送邮件的人名------------

# 例如，输入["结算款_1_周","结算款_2_朱朱"]，输出['周', '朱朱']
# def  getPersonName(att_name):
# 	person_name = []
# 	for name in att_name:
# 		person_name.append(name.split("_")[-1])
# 	return person_name

# person_name = getPersonName(att_name)

# -------------------根据附件名称中获得的人名，查找通讯录，找到对应的邮件地址---------------

def getRecvAddr(addrs,person_name):
	if not person_name in addrs:
		print("没有此人的邮箱地址!")
	return addrs[person_name]



# --------------------加载邮件内容-------------------------

def getMailContent(content_path):
	mail_content = ''

	if not os.path.exists(content_path):
		print("文件 content.txt 不存在")
		exit(0)

	with open(content_path,'r') as contentFile:
		contentLines = contentFile.readlines()
		if len(contentLines) < 1:
			print("no content in content.txt ")
			exit(0)
		mail_content = "".join(contentLines) # 将其+""转为字符串就好了
	return mail_content 


# --------------------添加附件-----------------------------------

def addAttch(attach_file):
	att = MIMEBase('application','octet-stream')  # 这两个参数不知道啥意思，二进制流文件
	att.set_payload(open(attach_file,'rb').read())
	# 此时的附件名称为****.xlsx，截取文件名
	att.add_header('Content-Disposition', 'attachment', filename=('gbk','', attach_file.split("\\")[-1]))
	encoders.encode_base64(att)
	return att



# ---------------------发送邮件-----------------------
def mailSend():
	smtp = smtplib.SMTP()   # 新建smtp对象
	smtp.connect(sender_host)
	smtp.login(sender_user, sender_pwd)

	for root,dirs,files in os.walk(attach_path):
		for attach_file in files:      # attach_file : ***_2_***.xlsx
			msg = MIMEMultipart('alternative')
			att_name = attach_file.split(".")[0]
			subject = att_name
			msg['Subject'] = subject   # 设置邮件主题
			person_name = subject.split("_")[-1]

			addrs = getAddrBook(addrBook)
			recv_addr = getRecvAddr(addrs,person_name)
			
			msg['From'] = formataddr([sender_name,sender_user]) # 设置发件人名称
			# msg['To'] = person_name # 设置收件人名称
			msg['To'] = formataddr([person_name,recv_addr]) # 设置收件人名称	
			mail_content = getMailContent(content_path)	
			msg.attach(MIMEText(mail_content))  # 正文  MIMEText(content,'plain','utf-8')
			attach_file = root+"\\"+attach_file
			att = addAttch(attach_file)
			msg.attach(att)  # 附件
			smtp.sendmail(sender_user, [recv_addr,], msg.as_string())  # smtp.sendmail(from_addr, to_addrs, msg)
			print("已发送： "+person_name+" <"+recv_addr+">")		
		smtp.quit()



if __name__ == '__main__':
	print("By 小周")
	mailSend()
	anykey = input("请按任意键退出程序：")
	if anykey:
		exit(5)



