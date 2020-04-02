# import main
from LayerWin import ask_run_win
from contact import contact_win
from readme import use_it
import wx
import time
import threading

class AskRun(ask_run_win):
	def __init__(self, parent=None, id=wx.ID_ANY, title='发票辅助工具'):
		ask_run_win.__init__(self, parent)

	def run_now( self, event ):
		self.Close()
		self.checked_it = True
		self.time_over = 0
		# 这里读取任务里第一条任务
		self.main_win.parse_task()

	def time_limit( self, event ):
		event.Skip()

	def close_calll_back( self, event ):
		event.Skip()

	def run_later( self, event ):
		self.Close()
		# 用于关闭窗口延迟计算
		self.time_over = 30
		self.retry = True

	def call(self,main_win):
		if hasattr(self,'check_layer') == False:
			self.main_win = main_win
			self.checked_it = False
			self.choose_run = False
			self.time_over = 10
			self.retry = False
			self.Show()
			self.check_layer = threading.Thread(target=self.run_check)
			self.cond = threading.Condition() # 锁
			self.check_layer.start()
			return "任务终端已接收"
		else:
			return "任务终端正在等待其它任务执行"	
				
	def run_check(self):
		while self.time_over > 0:
			self.m_staticText3.SetLabelText('%ss 后自动关闭窗口并运行' % self.time_over)
			time.sleep(1)
			self.time_over -= 1
		if self.retry == True:
			self.time_over = 10
			self.retry = False
			self.Show()
			return self.run_check()
		
		if self.checked_it == False:
			self.Close()
			self.checked_it = True
			# 设置不提醒弹框
			self.main_win.tip = False
			self.main_win.parse_task()
		delattr(self,'check_layer')
		print('线程结束',hasattr(self,'check_layer'))	

class Contact(contact_win):
	def __init__(self, parent=None):
		contact_win.__init__(self, parent)

	def call(self,main_win):
		self.main_win = main_win
		self.Show()
	def make_sure( self, event ):
		self.Close()

class Guide(use_it):
	def __init__(self, parent=None):
		use_it.__init__(self, parent)

	def call(self,main_win):
		self.main_win = main_win
		self.Show()

	def know( self, event ):
		self.Close()