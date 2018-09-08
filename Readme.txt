=================项目简介=================
针对上海速优装饰有限公司每月的结算款、预算款邮件发送需求的工具，
用于批量发送邮件，每个邮件包括含收件人名字的附件

项目开始于：2018年8月25号
开发人员：小周
-----------------V2.0-----------------

完成时间：2018-09-02

1、在V1.x版本基础上，初步建立GUI操作0界面，在GUI界面选择附件文件夹、邮箱通讯录文件和撰写邮件内容（纯文本）。

2、增加MyGUI.py文件，使用tkinter模块。

-----------------V1.0 V1.1----------------

完成时间：2018-08-29
1、功能：初步完成邮件发送，需要预先输入附件文件夹，和邮箱通讯录（CSV文件），并把邮箱内容存在TXT文件中；

2、使用smtp，email模块

2、包括：

sender_host  # 默认服务器地址及端口
sender_user  
sender_pwd = 
sender_name = u'上海周YJ公司'
attach_path = r'C:\Users\attchfile'   # 附件所在文件夹
attach_type = ".xlsx"      # 附件后缀名，即类型
addrBook = r'C:\邮箱联系人表单.csv'  # 邮件地址通讯录
content_path = r"C:\content.txt"   # 邮箱正文内容.txt

# 根据输入的CSV文件，获取通讯录人名和相应的邮箱地址---def getAddrBook(addrBook)
# 根据附件名称中获得的人名，查找通讯录，找到对应的邮件地址--def getRecvAddr(addrs,person_name):
# 加载邮件内容---def getMailContent(content_path):
# 添加附件--def addAttch(attach_file):
# 发送邮件---def mailSend():

V1.1中，只是将V1.0中多余的函数去掉的简洁版本，功能一样。

