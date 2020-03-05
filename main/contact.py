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
## Class contact_win
###########################################################################

class contact_win ( wx.Dialog ):

	def __init__( self, parent ):
		wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = u"合肥鑫山财务管理有限公司", pos = wx.DefaultPosition, size = wx.Size( 378,241 ), style = wx.DEFAULT_DIALOG_STYLE )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

		bSizer5 = wx.BoxSizer( wx.VERTICAL )

		sbSizer5 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"联系我们" ), wx.VERTICAL )

		sbSizer5.SetMinSize( wx.Size( 80,80 ) )
		self.m_staticText12 = wx.StaticText( sbSizer5.GetStaticBox(), wx.ID_ANY, u"联 系 人：王经理", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText12.Wrap( -1 )

		sbSizer5.Add( self.m_staticText12, 0, wx.ALL, 5 )

		self.m_staticText10 = wx.StaticText( sbSizer5.GetStaticBox(), wx.ID_ANY, u"联系地址：安徽省合肥市站西路宝文国际大厦702到707室", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText10.Wrap( -1 )

		self.m_staticText10.SetFont( wx.Font( wx.NORMAL_FONT.GetPointSize(), wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString ) )
		self.m_staticText10.SetForegroundColour( wx.Colour( 0, 0, 0 ) )

		sbSizer5.Add( self.m_staticText10, 0, wx.ALL, 5 )

		self.m_staticText13 = wx.StaticText( sbSizer5.GetStaticBox(), wx.ID_ANY, u"电子邮箱：2713913164@qq.com", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText13.Wrap( -1 )

		sbSizer5.Add( self.m_staticText13, 0, wx.ALL, 5 )

		self.m_staticText14 = wx.StaticText( sbSizer5.GetStaticBox(), wx.ID_ANY, u"联系电话：13696546411", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText14.Wrap( -1 )

		sbSizer5.Add( self.m_staticText14, 0, wx.ALL, 5 )

		self.m_staticText15 = wx.StaticText( sbSizer5.GetStaticBox(), wx.ID_ANY, u"Q       Q：646591494", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText15.Wrap( -1 )

		sbSizer5.Add( self.m_staticText15, 0, wx.ALL, 5 )


		bSizer5.Add( sbSizer5, 1, wx.EXPAND, 5 )

		bSizer11 = wx.BoxSizer( wx.VERTICAL )


		bSizer11.Add( ( 0, 0), 1, wx.EXPAND, 5 )

		self.m_button4 = wx.Button( self, wx.ID_ANY, u"确定", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer11.Add( self.m_button4, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


		bSizer5.Add( bSizer11, 1, wx.EXPAND, 5 )


		self.SetSizer( bSizer5 )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.m_button4.Bind( wx.EVT_BUTTON, self.make_sure )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def make_sure( self, event ):
		event.Skip()


