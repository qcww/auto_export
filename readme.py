# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version 3.9.0 Dec  4 2019)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
## Class MyDialog2
###########################################################################

class use_it ( wx.Dialog ):

	def __init__( self, parent ):
		wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 513,525 ), style = wx.DEFAULT_DIALOG_STYLE )

		self.SetSizeHints( wx.Size( 513,525 ), wx.Size( 513,525 ) )

		bSizer4 = wx.BoxSizer( wx.VERTICAL )

		bSizer5 = wx.BoxSizer( wx.VERTICAL )

		self.m_staticText13 = wx.StaticText( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( -1,-1 ), 0 )
		self.m_staticText13.Wrap( -1 )

		bSizer5.Add( self.m_staticText13, 0, wx.ALL, 5 )

		self.m_staticText6 = wx.StaticText( self, wx.ID_ANY, u"鑫山财务税控发票导出辅助软件使用说明", wx.DefaultPosition, wx.Size( -1,-1 ), 0 )
		self.m_staticText6.Wrap( -1 )

		self.m_staticText6.SetFont( wx.Font( 13, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, wx.EmptyString ) )

		bSizer5.Add( self.m_staticText6, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


		bSizer4.Add( bSizer5, 1, wx.EXPAND, 5 )

		sbSizer2 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"软件介绍" ), wx.VERTICAL )

		self.m_staticText8 = wx.StaticText( sbSizer2.GetStaticBox(), wx.ID_ANY, u"本软件旨在辅助开票软件及相关报税所需软件，导出并自动上传会计报税做账所", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText8.Wrap( -1 )

		sbSizer2.Add( self.m_staticText8, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText9 = wx.StaticText( sbSizer2.GetStaticBox(), wx.ID_ANY, u"需资料，仅供内部使用，严禁外传。", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText9.Wrap( -1 )

		sbSizer2.Add( self.m_staticText9, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


		bSizer4.Add( sbSizer2, 1, wx.EXPAND, 5 )

		sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"软件安装" ), wx.VERTICAL )

		self.m_staticText14 = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"本软件无需安装，建议放置于非系统盘下，使用软件前请确保电脑系统win7及", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText14.Wrap( -1 )

		sbSizer3.Add( self.m_staticText14, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText15 = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"以上，已安装开票软件与相关驱动程序。首次运行辅助软件前请打开开票软件", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText15.Wrap( -1 )

		sbSizer3.Add( self.m_staticText15, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText151 = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"并插入税控盘。导出进项票需要统一安装指定谷歌版本浏览器。", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText151.Wrap( -1 )

		sbSizer3.Add( self.m_staticText151, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText17 = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"软件支持开机自启动(需首次打开开票软件并运行本软件)、自动处理消息", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText17.Wrap( -1 )

		sbSizer3.Add( self.m_staticText17, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


		bSizer4.Add( sbSizer3, 1, wx.EXPAND, 5 )

		sbSizer4 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"注意事项" ), wx.VERTICAL )

		self.m_staticText18 = wx.StaticText( sbSizer4.GetStaticBox(), wx.ID_ANY, u"本软件仅用于辅助上传会计做账报税所需资料，无毒无广告，尊重客户隐私，", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText18.Wrap( -1 )

		self.m_staticText18.SetMinSize( wx.Size( 420,-1 ) )

		sbSizer4.Add( self.m_staticText18, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText19 = wx.StaticText( sbSizer4.GetStaticBox(), wx.ID_ANY, u"请放心使用。如遇杀毒软件提示木马病毒请添加到信任文件后使用。", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText19.Wrap( -1 )

		sbSizer4.Add( self.m_staticText19, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText20 = wx.StaticText( sbSizer4.GetStaticBox(), wx.ID_ANY, u"不建议同时插入多个税控盘设备时使用本软件，可能导致上传错误或失败。辅助", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText20.Wrap( -1 )

		sbSizer4.Add( self.m_staticText20, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		self.m_staticText21 = wx.StaticText( sbSizer4.GetStaticBox(), wx.ID_ANY, u"软件运行时建议勿操作干扰以免影响资料上传", wx.DefaultPosition, wx.Size( 420,-1 ), 0 )
		self.m_staticText21.Wrap( -1 )

		sbSizer4.Add( self.m_staticText21, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


		bSizer4.Add( sbSizer4, 2, wx.EXPAND, 5 )

		self.m_button2 = wx.Button( self, wx.ID_ANY, u"确定", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer4.Add( self.m_button2, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


		self.SetSizer( bSizer4 )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.m_button2.Bind( wx.EVT_BUTTON, self.know )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def know( self, event ):
		event.Skip()


