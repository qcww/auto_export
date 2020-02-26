# -*- coding: utf-8 -*- 

import wx
import ui
import time
import configparser
import AutoExport
import datetime
import requests
import re
import os
 
class Ep(ui.MyFrame1):
    def start(self):
        self.config_file = './config/config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')
        self.reset_date()
        self.set_status('已准备')
        self.set_dw_link()
        self.add_log('辅助工具已打开')
        self.app_status()

    # 软件下载链接
    def set_dw_link(self):
        self.m_hyperlink2.SetURL(self.config['app']['download_ht']) 

    # 重新设置默认时间
    def reset_date(self):
        for y in range(6):
            strTime = time.strftime("%Y",  time.localtime(time.time()-86400*365*y))
            self.m_choice3.SetString(y,strTime+'年')

        today = datetime.date.today()
        first = today.replace(day=1)
        lastMonth = first - datetime.timedelta(days=1)
        ym = lastMonth.strftime("%Y-%m")
        ym_split = ym.split('-')
        now = time.strftime("%Y-%m-01",  time.localtime())
        now_split = now.split('-')
        dec_year = int(now_split[0]) - int(ym_split[0])
        dec_month = int(ym_split[1])
        self.m_choice3.SetSelection(dec_year)
        self.m_choice4.SetSelection(dec_month - 1)

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
        # auto_app = AutoExport.Export()
        # # 最小化本应用窗口
        # auto_app.min_self()
        # auto_app.do_ready()
        # # 修复数据
        # auto_app.fix_data(ym)
        # ret = auto_app.dw_excel(ym)
        # auto_app.min_app()
        # 导出上传数据
        # self.add_log(ret['msg'])
        ret = {"code":200}
        if ret['code'] == 200:
            self.add_log("下载成功，正在上传")
            sales_name = self.config['upload']['sales_name'] + time.strftime("%Y%m%d",  time.localtime()) + self.config['upload']['sales_file_ext']
            sales_upload_link = self.config['upload']['sales_upload_link']
            ym_split = ym.split('-')
            post_data = {"credit_code":self.credit_code,"period":ym_split[0]+ym_split[1],"submit":"1","tax_import":"1"}

            up_res = self.upload_file(sales_name,sales_upload_link,post_data)
            if up_res['code'] == 0:
                dlg = wx.MessageDialog(None, u"上传成功", u"提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
            else:
                dlg = wx.MessageDialog(None, up_res['text'] + "是否重新运行", u"发生错误", wx.YES_NO | wx.STAY_ON_TOP | wx.ICON_EXCLAMATION)

            self.add_log(up_res['text'])
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


    # 上传文件
    def upload_file(self,file_name,upload_link,data):
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        file_name = './exp_file/'+file_name
        if os.path.exists(file_name) == False:
            return {"code":500,"text":"导出数据失败"}
        files = {'upfile': open(file_name, 'rb')}
        res = requests.post(upload_link, data=data, files=files, headers=headers)
        print(res.content)
        return res.json()
        # try:
        #     res = requests.post(upload_link, data=data, files=files, headers=headers)
        #     return res.json()
        # except:
        #     return {"code":500,"text":"网络异常，上传失败"}

    # 添加日志
    def add_log(self,text):
        if text == "":
            return False
        self.m_listBox3.InsertItems(["%s %s" % (time.strftime("%M:%S", time.localtime()),text)],0)

    # 开票软件运行状态
    def app_status(self):
        self.set_config('log','last_run_time',time.strftime("%Y-%m-%d %H:%M", time.localtime()))
        auto_app = AutoExport.Export()
        auto_app.do_ready()
        self.credit_code,pname = auto_app.user_info()
        if pname != "":
            self.set_status('开票软件已运行 | ' + pname)
        self.m_staticText5.SetLabelText('系统检测到三月份数据未导出')
        # self.m_staticText5.SetForegroundColour( wx.Colour( 76, 146, 76 ) )

    # 获取选择的时间
    def get_y_m(self):
        ch_year = self.m_choice3.GetSelection()
        selec_year = time.strftime("%Y",  time.localtime(time.time()-86400*365*ch_year))
        ch_month = self.m_choice4.GetSelection()
        return "%4d-%02d-01" % (int(selec_year), int(ch_month)+1)

    # 写入配置
    def set_config(self,section,key,value):
        set_val = str(value)
        self.config.set(section,key,re.sub(r'%','#53',set_val.strip('|')))
        self.config.write(open(self.config_file,'w',encoding='utf-8'))

app = wx.App()
frame = Ep(None)
frame.start()
frame.Show()
app.MainLoop()