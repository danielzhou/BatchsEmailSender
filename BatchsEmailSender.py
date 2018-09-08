
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
#

# --------------------发送服务器配置---------------

# # addrBook = r'C:\Users\小周\Desktop\批量发送邮件\邮箱联系人表单.csv'  # 邮件地址通讯录)
# content_path = r"C:\Users\小周\Desktop\批量发送邮件\结算款邮件内容.txt"   # 邮箱正文内容.txt

"""docstring for BatchsMailSender"""
sender_host = 'smtp.163.com:25'  # 默认服务器地址及端口
sender_user = 'laoxxx@163.com'
sender_pwd = '123456'
sender_name = u'上海周YJ公司'
# self.attach_type = ".xlsx"


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

# -------------------根据附件名称中获得的人名，查找通讯录，找到对应的邮件地址---------------

def getRecvAddr(addrs,person_name):
	if not person_name in addrs:
		print("没有<"+person_name+">的邮箱地址! 请在联系人中添加此人邮箱后重试。")
		anykey = input("请按回车键（Enter）退出程序：")
		if anykey:
			sys.exit(0)
	# try:
	return addrs[person_name]
	# except KeyError:
	# 	print("通讯录中无此人："+person_name)
	# 	# raise SystemExit

# --------------------添加附件-----------------------------------

def addAttch(attach_file):
	att = MIMEBase('application','octet-stream')  # 这两个参数不知道啥意思，二进制流文件
	att.set_payload(open(attach_file,'rb').read())
	# 此时的附件名称为****.xlsx，截取文件名
	att.add_header('Content-Disposition', 'attachment', filename=('gbk','', attach_file.split("\\")[-1]))
	encoders.encode_base64(att)
	return att



# ---------------------发送邮件-----------------------
def mailSend(attach_path,bookFile,mail_content):
	smtp = smtplib.SMTP()   # 新建smtp对象
	smtp.connect(sender_host)
	smtp.login(sender_user, sender_pwd)
	addrBook = bookFile
	# attach_path = attach_path
	# mail_content = mail_content
	addrs = getAddrBook(addrBook)

	for root,dirs,files in os.walk(attach_path):
		for attach_file in files:      # attach_file : ***_2_***.xlsx
			msg = MIMEMultipart('alternative')
			att_name = attach_file.split(".")[0]
			subject = att_name
			msg['Subject'] = subject   # 设置邮件主题
			person_name = subject.split("_")[-1]
			recv_addr = getRecvAddr(addrs,person_name)
			msg['From'] = formataddr([sender_name,sender_user]) # 设置发件人名称
			# msg['To'] = person_name # 设置收件人名称
			msg['To'] = formataddr([person_name,recv_addr]) # 设置收件人名称
			# mail_content = getMailContent(content_path)
			msg.attach(MIMEText(mail_content))  # 正文  MIMEText(content,'plain','utf-8')
			attach_file = root+"\\"+attach_file
			att = addAttch(attach_file)
			msg.attach(att)  # 附件
			smtp.sendmail(sender_user, [recv_addr,], msg.as_string())  # smtp.sendmail(from_addr, to_addrs, msg)
			print("已发送： "+person_name+" <"+recv_addr+">")
		smtp.quit()
	sys.exit("Have a nice day !")



