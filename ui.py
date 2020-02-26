# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import wx.adv

###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1 ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"鑫山财务-开票辅助工具", pos = wx.DefaultPosition, size = wx.Size( 450,365 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.Size( 450,365 ), wx.Size( 450,365 ) )
		self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_BTNFACE ) )
		
		self.m_menubar1 = wx.MenuBar( 0 )
		self.m_menu6 = wx.Menu()
		self.m_menuItem2 = wx.MenuItem( self.m_menu6, wx.ID_ANY, u"环境检测", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu6.AppendItem( self.m_menuItem2 )
		
		self.m_menubar1.Append( self.m_menu6, u"操作" ) 
		
		self.m_menu7 = wx.Menu()
		self.m_menuItem3 = wx.MenuItem( self.m_menu7, wx.ID_ANY, u"软件使用", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu7.AppendItem( self.m_menuItem3 )
		
		self.m_menuItem5 = wx.MenuItem( self.m_menu7, wx.ID_ANY, u"软件版本", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu7.AppendItem( self.m_menuItem5 )
		
		self.m_menubar1.Append( self.m_menu7, u"帮助" ) 
		
		self.m_menu10 = wx.Menu()
		self.m_menuItem6 = wx.MenuItem( self.m_menu10, wx.ID_ANY, u"联系地址", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu10.AppendItem( self.m_menuItem6 )
		
		self.m_menubar1.Append( self.m_menu10, u"联系我们" ) 
		
		self.SetMenuBar( self.m_menubar1 )
		
		gSizer2 = wx.GridSizer( 0, 2, 0, 0 )
		
		sbSizer7 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"软件日志" ), wx.VERTICAL )
		
		m_listBox3Choices = []
		self.m_listBox3 = wx.ListBox( sbSizer7.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.Size( 250,340 ), m_listBox3Choices, 0 )
		self.m_listBox3.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		sbSizer7.Add( self.m_listBox3, 0, wx.ALL, 5 )
		
		
		gSizer2.Add( sbSizer7, 1, wx.EXPAND, 5 )
		
		bSizer8 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer11 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"辅助工具" ), wx.VERTICAL )
		
		self.m_staticText2 = wx.StaticText( sbSizer11.GetStaticBox(), wx.ID_ANY, u"提示：切换开票软件后重启软件生效", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText2.Wrap( -1 )
		self.m_staticText2.SetForegroundColour( wx.Colour( 239, 82, 78 ) )
		self.m_staticText2.SetBackgroundColour( wx.Colour( 240, 240, 240 ) )
		
		sbSizer11.Add( self.m_staticText2, 0, wx.ALL, 5 )
		
		self.m_staticText1 = wx.StaticText( sbSizer11.GetStaticBox(), wx.ID_ANY, u"开票软件", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText1.Wrap( -1 )
		self.m_staticText1.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INACTIVECAPTIONTEXT ) )
		
		sbSizer11.Add( self.m_staticText1, 0, wx.ALL, 5 )
		
		bSizer3 = wx.BoxSizer( wx.VERTICAL )
		
		gSizer4 = wx.GridSizer( 0, 2, 0, 0 )
		
		self.m_radioBtn6 = wx.RadioButton( sbSizer11.GetStaticBox(), wx.ID_ANY, u"航天", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_radioBtn6.SetValue( True ) 
		self.m_radioBtn6.SetFont( wx.Font( 9, 70, 90, 90, False, "宋体" ) )
		self.m_radioBtn6.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		self.m_radioBtn6.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INFOTEXT ) )
		
		gSizer4.Add( self.m_radioBtn6, 0, wx.ALL, 5 )
		
		self.m_hyperlink2 = wx.adv.HyperlinkCtrl( sbSizer11.GetStaticBox(), wx.ID_ANY, u"下载", wx.EmptyString, wx.DefaultPosition, wx.DefaultSize)
		gSizer4.Add( self.m_hyperlink2, 0, wx.ALL, 5 )
		
		
		bSizer3.Add( gSizer4, 1, wx.EXPAND, 5 )
		
		gSizer3 = wx.GridSizer( 0, 2, 0, 0 )
		
		self.m_radioBtn7 = wx.RadioButton( sbSizer11.GetStaticBox(), wx.ID_ANY, u"百旺", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer3.Add( self.m_radioBtn7, 0, wx.ALL, 5 )
		
		self.m_hyperlink3 = wx.adv.HyperlinkCtrl( sbSizer11.GetStaticBox(), wx.ID_ANY, u"下载", wx.EmptyString, wx.DefaultPosition, wx.DefaultSize )
		gSizer3.Add( self.m_hyperlink3, 0, wx.ALL, 5 )
		
		
		bSizer3.Add( gSizer3, 1, wx.EXPAND, 5 )
		
		
		sbSizer11.Add( bSizer3, 1, wx.EXPAND, 5 )
		
		
		bSizer8.Add( sbSizer11, 1, wx.EXPAND, 5 )
		
		sbSizer12 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"操作" ), wx.VERTICAL )
		
		bSizer9 = wx.BoxSizer( wx.VERTICAL )
		
		fgSizer3 = wx.FlexGridSizer( 0, 3, 0, 0 )
		fgSizer3.SetFlexibleDirection( wx.BOTH )
		fgSizer3.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.m_staticText4 = wx.StaticText( sbSizer12.GetStaticBox(), wx.ID_ANY, u"选择时间", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText4.Wrap( 1 )
		fgSizer3.Add( self.m_staticText4, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		m_choice3Choices = [ u"请选择", u"请选择", u"请选择", u"请选择", u"请选择", u"请选择" ]
		self.m_choice3 = wx.Choice( sbSizer12.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice3Choices, wx.CB_SORT )
		self.m_choice3.SetSelection( 1 )
		fgSizer3.Add( self.m_choice3, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		m_choice4Choices = [ u"1月", u"2月", u"3月", u"4月", u"5月", u"6月", u"7月", u"8月", u"9月", u"10月", u"11月", u"12月" ]
		self.m_choice4 = wx.Choice( sbSizer12.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice4Choices, 0 )
		self.m_choice4.SetSelection( 0 )
		fgSizer3.Add( self.m_choice4, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		bSizer9.Add( fgSizer3, 1, wx.EXPAND, 5 )
		
		self.m_staticText5 = wx.StaticText( sbSizer12.GetStaticBox(), wx.ID_ANY, u"数据同步中...", wx.Point( -1,-1 ), wx.DefaultSize, wx.ALIGN_CENTRE )
		self.m_staticText5.Wrap( -1 )
		self.m_staticText5.SetForegroundColour( wx.Colour( 239, 82, 78 ) )
		self.m_staticText5.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		bSizer9.Add( self.m_staticText5, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		sbSizer12.Add( bSizer9, 1, wx.EXPAND, 5 )
		
		gSizer6 = wx.GridSizer( 0, 2, 0, 0 )
		
		self.m_button6 = wx.Button( sbSizer12.GetStaticBox(), wx.ID_ANY, u"销项票导出", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer6.Add( self.m_button6, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.TOP|wx.BOTTOM, 5 )
		
		self.m_button7 = wx.Button( sbSizer12.GetStaticBox(), wx.ID_ANY, u"进项票导出", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer6.Add( self.m_button7, 0, wx.TOP|wx.BOTTOM, 5 )
		
		
		sbSizer12.Add( gSizer6, 1, wx.EXPAND, 5 )
		
		
		bSizer8.Add( sbSizer12, 1, wx.EXPAND, 5 )
		
		
		gSizer2.Add( bSizer8, 1, wx.EXPAND, 5 )
		
		
		self.SetSizer( gSizer2 )
		self.Layout()
		self.m_statusBar1 = self.CreateStatusBar( 1, 0, wx.ID_ANY )
		self.m_statusBar1.SetForegroundColour( wx.Colour( 255, 0, 0 ) )
		self.m_statusBar1.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.Bind( wx.EVT_MENU, self.contact_us, id = self.m_menuItem6.GetId() )
		self.m_radioBtn6.Bind( wx.EVT_RADIOBUTTON, self.toggle )
		self.m_radioBtn7.Bind( wx.EVT_RADIOBUTTON, self.toggle )
		self.m_choice3.Bind( wx.EVT_CHOICE, self.select_year )
		self.m_choice4.Bind( wx.EVT_CHOICE, self.select_month )
		self.m_button6.Bind( wx.EVT_BUTTON, self.export_sales )
		self.m_button6.Bind( wx.EVT_LEAVE_WINDOW, self.m_button6OnLeaveWindow )
		self.m_button6.Bind( wx.EVT_LEFT_DCLICK, self.m_button6OnLeftDClick )
		self.m_button7.Bind( wx.EVT_KEY_DOWN, self.export_purchase )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def contact_us( self, event ):
		event.Skip()
	
	def toggle( self, event ):
		event.Skip()
	
	
	def select_year( self, event ):
		event.Skip()
	
	def select_month( self, event ):
		event.Skip()
	
	def export_sales( self, event ):
		event.Skip()
	
	def m_button6OnKeyUp( self, event ):
		event.Skip()
	
	def m_button6OnLeaveWindow( self, event ):
		event.Skip()
	
	def m_button6OnLeftDClick( self, event ):
		event.Skip()
	
	def export_purchase( self, event ):
		event.Skip()
	

