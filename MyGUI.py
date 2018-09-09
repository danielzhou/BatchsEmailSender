# -*- coding:utf-8 -*-

# ==================BatchMailSender V1.3 ===========================

import tkinter
from tkinter.constants import *
from tkinter import *
# from tkinter import ttk
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText
import tkinter.messagebox
from tkinter import END
import time
import BatchsEmailSender



class MyGUI():
	"""@brief:使用GUI界面来选择邮箱附件所在的文件夹和邮箱通讯录
	"""

	# 构造函数
	def __init__(self,init_window_name):
		self.init_window_name = init_window_name

	# 初始化窗口
	def set_init_window(self):
		self.init_window_name.title("批量发送邮件-for 速优 -- By 小周 ")  # 设置标题
		self.init_window_name.geometry('600x480+300+200')  # 设置尺寸
		self.init_window_name.attributes('-alpha',1)  # 属性？

		# 设置标签
		self.attach_label = Label(self.init_window_name,text="请选择附件所在文件夹：",bg='lightblue')
		self.attach_label.place(x=50,y=53,width=180,height=30)
		self.book_label = Label(self.init_window_name,text="请选择邮箱地址通讯录（CSV文件）：",justify='left',bg='lightblue')
		self.book_label.place(x=50,y=125,width=180,height=30)
		self.info_label = Label(self.init_window_name,text="请输入邮件正文内容：",bg='lightblue')
		self.info_label.place(x=50,y=190,width=180,height=30)


		# 设置输入框应该与标签同一竖直--x相同
		# self.input_1 = StringVar()
		# self.attach_dir = Entry(self.init_window_name,textvariable=input_1)
		# self.attach_dir.place(x=50,y=88,width=300,height=30)
		# self.input_2 = StringVar()
		# self.book_dir = Entry(self.init_window_name,textvariable=input_2)
		# self.book_dir.place(x=50,y=158,width=300,height=30)

		self.attach_dir = Label(self.init_window_name,text="",relief='groove',justify='left')
		self.attach_dir.place(x=50,y=88,width=400,height=30)
		self.book_dir = Label(self.init_window_name,text="",relief='groove',justify='left')
		self.book_dir.place(x=50,y=158,width=400,height=30)


		# 设置按钮--应该与输入框同一水平--y相同
		self.attach_button = Button(self.init_window_name,text="选择文件夹",bg='lightblue',width=15,command=self.selectAttachDir)
		self.attach_button.place(x=460,y=88,width=80,height=30)
		self.book_button = Button(self.init_window_name,text="选择通讯录",bg='lightblue',width=15,command=self.selectBookFile)
		self.book_button.place(x=460,y=158,width=80,height=30)

		# 设置邮件内容
		self.info_text = ScrolledText(self.init_window_name)
		self.info_text.place(x=50,y=220,width=400,height=180)
		self.info_text.insert(INSERT,"")
		self.info_text.insert(END,"")
		# # print("邮件内容："+str(self.info_text.get(1.0,END)))
		# print(type(self.info_text.get(END)))
		# print(self.info_text.get(0.0,END))

		# 设置确定/取消按钮
		self.finish_button = Button(self.init_window_name,text="确定",bg='lightblue',width=15,command=self.start_send_mail) #　command=self.test
		self.finish_button.place(x=460,y=430,width=80,height=30)  #
		self.cancel_button = Button(self.init_window_name,text="取消",bg='lightblue',width=15,command=self.init_window_name.quit)
		self.cancel_button.place(x=370,y=430,width=80,height=30)  #




	# 选择附件目录
	def selectAttachDir(self):

		self.attachDir = askdirectory()
		# print(type(self.attachDir))   # <class 'str'>
		# self.input_1.set(attachDir)
		self.attach_dir.config(text=self.attachDir)  #　将结果显示在label上

	# 选择通讯录文件
	def selectBookFile(self):
		self.bookFile = askopenfilename()
		# print(type(self.bookFile))   # <class 'str'>
		# self.input_2.set(bookFile)
		self.book_dir.config(text=self.bookFile)

	# 测试如何获取 ScrolledText 中的文本内容
	def  test(self):
		print(type(self.info_text.get(END)))    # str格式
		print(self.info_text.get(0.0,END))   # 获取ScrolledText中的文本内容-正文

		# print(self.attach_dir.get())   # 'Label' object has no attribute 'get'
		# print(self.book_dir.get())
		print(self.attach_dir['text'])    # 获取label中的文本内容-目录
		print(self.book_dir['text'])  # 获取label中的文本内容-通讯录
		self.init_window_name.quit

	# # 开始发送邮件--attach_path,bookFile,mail_content
	def start_send_mail(self):

		# attach_path = self.attach_dir['text']
		# bookFile = self.book_dir['text']
		# mail_content = self.info_text.get(0.0,END)
		# BatchsMailSender.mailSend(attach_path,bookFile,mail_content)

		BatchsEmailSender.mailSend(self.attach_dir['text'],self.book_dir['text'],self.info_text.get(0.0,END))

		self.init_window_name.quit




def gui_start():
	init_window = tkinter.Tk()
	main_gui = MyGUI(init_window)
	main_gui.set_init_window()
	init_window.mainloop()

# gui_start()

# 判断是否过试用期
def if_expired():
	end_time = '2018-11-26 10:00:00'  # 设置截止时间
	end_time = time.strptime(end_time,'%Y-%m-%d %H:%M:%S')
	end_time = int(time.mktime(end_time))

	cur_time = int(time.time()) # 获取当前时间

	left_time = end_time - cur_time   # 剩余时间

	if left_time <= 0:
		warning_window = tkinter.Tk()
		warning_window.title("批量发送邮件-for 速优 -- By 小周 ")
		warning_window.geometry('600x480+300+200')  # 设置尺寸
		warning_window.attributes('-alpha',1)  # 属性
		tkinter.messagebox.showwarning(title="温馨提示",message="试用期已过，如需续用，请联系作者邮箱：\n小周 <1358304569@qq.com>")
		warning_window.mainloop()
		sys.exit()
	else:
		pass


if __name__ == '__main__':
	# 判断是否到期
	if_expired()
	# 正常运行
	gui_start()


















