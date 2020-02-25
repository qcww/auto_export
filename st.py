# -*- coding: utf-8 -*- 

import wx
import ui
import time
import configparser
import AutoExport
 
class Ep(ui.MyFrame1):
    def start(self):
        self.config_file = './config/config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')
        self.reset_date()
        self.set_status('已准备')
        self.set_dw_link()
        self.add_log('辅助工具已打开')
        # self.app_status()

    # 软件下载链接
    def set_dw_link(self):
        self.m_hyperlink2.SetURL(self.config['app']['download_ht']) 

    # 重新设置默认时间
    def reset_date(self):
        for y in range(2):
            strTime = time.strftime("%Y",  time.localtime(time.time()-86400*365*y))
            self.m_choice3.SetString(y,strTime+'年')

    # 界面状态栏状态
    def set_status(self,text):
        self.m_statusBar1.SetStatusText(' %s ' % text)

    # 导出销项票
    def export_sales(self,event):
        # dlg = wx.MessageDialog(None, "是否重新运行", u"发生错误", wx.YES_NO | wx.STAY_ON_TOP | wx.ICON_EXCLAMATION)
        # if dlg.ShowModal() == wx.ID_YES:
        #     self.Close(True)
        # return False
        ym = self.get_y_m()
        auto_app = AutoExport.Export()
        # 最小化本应用窗口
        auto_app.min_self()
        ret = auto_app.dw_excel(ym)
        self.add_log(ret['msg'])

        if ret['code'] == 200:
            dlg = wx.MessageDialog(None, u"正在上传数据，请不要关闭软件", u"操作成功", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)
        elif ret['code'] == 404:
            dlg = wx.MessageDialog(None, ret['msg'], u"提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)
        else:
            dlg = wx.MessageDialog(None, ret['msg'] + "是否重新运行", u"发生错误", wx.YES_NO | wx.STAY_ON_TOP | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)

        dlg.Destroy()

    # 添加日志
    def add_log(self,text):
        if text == "":
            return False
        self.m_listBox3.InsertItems(["%s %s" % (time.strftime("%M:%S", time.localtime()),text)],0)

    # 开票软件运行状态
    def app_status(self):
        auto_app = AutoExport.Export()
        self.uid,pname = auto_app.user_info()
        if pname != "":
            self.set_status('开票软件已运行 | ' + pname)

    # 获取选择的时间
    def get_y_m(self):
        ch_year = self.m_choice3.GetSelection()
        selec_year = time.strftime("%Y",  time.localtime(time.time()-86400*365*ch_year))
        ch_month = self.m_choice4.GetSelection()
        return "%4d-%02d-01" % (int(selec_year), int(ch_month)+1)

app = wx.App()
frame = Ep(None)
frame.start()
frame.Show()
app.MainLoop()