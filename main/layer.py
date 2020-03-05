# import main
from LayerWin import ask_run_win
from contact import contact_win
from readme import use_it
import wx
import time
import threading

class AskRun(ask_run_win):
	def __init__(self, parent=None, id=wx.ID_ANY, title='鑫山财务-开票辅助工具'):
		ask_run_win.__init__(self, parent)

	def run_now( self, event ):
		self.Close()
		self.checked = True
		# 这里读取任务里第一条任务
		self.main.parse_task()

	def time_limit( self, event ):
		event.Skip()

	def close_calll_back( self, event ):
		self.time_over = 15
		event.Skip()

	def run_later( self, event ):
		self.Close()
		time.sleep(10)
		return self.call(self.main)


	def call(self,main):
		self.main = main
		self.checked = False
		self.time_over = 15
		
		self.Show()
		if hasattr(self,'check_layer') == False:
			self.check_layer = threading.Thread(target=self.run_check)
			self.cond = threading.Condition() # 锁
			self.check_layer.start()
				
	def run_check(self):
		while self.time_over > 0:
			self.m_staticText3.SetLabelText('%ss 后自动关闭窗口并运行' % self.time_over)
			time.sleep(1)
			print(self.checked)
			self.time_over -= 1
		self.Close()
		if self.checked == False:
			self.checked = True
			self.main.parse_task()

class Contact(contact_win):
	def __init__(self, parent=None):
		contact_win.__init__(self, parent)

	def call(self,main):
		self.main = main
		self.Show()
	def make_sure( self, event ):
		self.Close()

class Guide(use_it):
	def __init__(self, parent=None):
		use_it.__init__(self, parent)

	def call(self,main):
		self.main = main
		self.Show()

	def know( self, event ):
		self.Close()