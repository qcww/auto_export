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
import sys
import regedit
import win32com.client
import websocket
import threading
import json
import UsbCheck
import ContactUs

class TaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        wx.adv.TaskBarIcon.__init__(self)
        self.frame = frame
        self.SetIcon(wx.Icon(name='mondrian.ico', type=wx.BITMAP_TYPE_ICO), '鑫山财务-开票辅助工具')
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self.OnTaskBarLeftDClick)
 
    def OnTaskBarLeftDClick(self, event):
        if self.frame.IsIconized():
           self.frame.Iconize(False)
        if not self.frame.IsShown():
           self.frame.Show(True)
        self.frame.Raise()
 
    def OnClose(self, event):
        self.frame.Destroy()
        self.Destroy()

    def OnConfig(self, event):
        pass

    def OnAbout(self, event):
        pass

    # Menu数据
    def setMenuItemData(self):
        # return (("关闭", self.OnClose))
        return (("配置", self.OnConfig), ("使用", self.OnAbout), ("退出", self.OnClose))
    
    # 创建菜单
    def CreatePopupMenu(self):
        menu = wx.Menu()
        for itemName, itemHandler in self.setMenuItemData():
            if not itemName:    # itemName为空就添加分隔符
                menu.AppendSeparator()
                continue
            menuItem = wx.MenuItem(None, wx.ID_ANY, text=itemName, kind=wx.ITEM_NORMAL) # 创建菜单项
            menu.Append(menuItem)                                                   # 将菜单项添加到菜单
            self.Bind(wx.EVT_MENU, itemHandler, menuItem)
        return menu



class Ep(ui.MyFrame1):
    def __init__(
            self, parent=None, id=wx.ID_ANY, title='鑫山财务-开票辅助工具', pos=wx.DefaultPosition,
            size=wx.DefaultSize, style=wx.DEFAULT_FRAME_STYLE
            ):
        ui.MyFrame1.__init__(self, parent)  
 
        # create a welcome screen
        # screen = wx.Image(self.screenIm).ConvertToBitmap()
        # wx.SplashScreen(screen, wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT,1000, None, -1)
        # wx.Yield()
       
        self.SetIcon(wx.Icon('mondrian.ico', wx.BITMAP_TYPE_ICO))
        panel = wx.Panel(self, wx.ID_ANY)
        button = wx.Button(panel, wx.ID_ANY, 'Hide Frame', pos=(60, 60))
       
        sizer = wx.BoxSizer()
        sizer.Add(button, 0)
        panel.SetSizer(sizer)
        self.taskBarIcon = TaskBarIcon(self)        
       
        # bind event
        self.Bind(wx.EVT_BUTTON, self.OnHide, button)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.Bind(wx.EVT_ICONIZE, self.OnIconfiy) # 最小化事件绑定

    def OnHide(self, event):
        self.Hide()

    def OnIconfiy(self, event):
        # wx.MessageBox('Frame has been iconized!', 'Prompt')
        event.Skip()

    def OnClose(self, event):
        self.Hide()

    def contact_us(self,event):
        print('联系我们')
        contact = ContactUs.MyPanel1(self, parent)

    def start(self):
        # 默认消息弹框提醒
        self.tip = True
        self.running_index = False
        self.config_file = "./config/config.ini"
        if os.path.exists(self.config_file) == False:
            self.config_file = regedit.get_client_path()+"\\config\\config.ini"
        else:
            regedit.set_client_path()

        self.uid,first_run = regedit.get_pc_id()
        if first_run == True:
            self.post_link(self.config['client']['update_client_url',{"invoice_export_client_id":self.uid}])
        try:
            self.config = configparser.ConfigParser()
            self.config.read(self.config_file,encoding='utf-8')
            self.host = self.config['link']['host']
        except:
            dlg = wx.MessageDialog(None, "检测到配置文件丢失", u"错误提示(鑫山财务-开票辅助工具)", wx.OK | wx.STAY_ON_TOP)
            if dlg.ShowModal() == wx.ID_YES:
                return False
            pass

        self.reset_date()
        self.set_status('已准备')

        if self.check_user_service(True) == False:
            return False

        self.set_dw_link()
        self.add_log('辅助工具已打开')
        # 设置开机启动
        main_app = regedit.get_client_path() + self.config['client']['name'] + " /autorun"
        if regedit.add_auto_run(main_app) == False:
            print('设置开机启动失败')
        self.check_usb_action()
        # self.app_status()
        self.connect_service()
        return True

    def check_usb_action(self):
        # self.run_status = bool(1 - self.run_status)
        if hasattr(self,'check_usb') == False:
            self.check_usb = threading.Thread(target=self.check_usb_staus)
            self.cond = threading.Condition() # 锁
            self.check_usb.start()

    def check_usb_staus(self,first = True):
        if self.config['app']['default'] == '6':
            check_usb = UsbCheck.check_usb(self.config['app']['usb_ht'])
            if check_usb == True:
                self.m_staticText41.SetForegroundColour( wx.Colour( 0, 128, 0 ) )
                self.m_staticText41.SetLabelText('金税盘已插入')
                self.usb_insert = True
            else:
                self.m_staticText41.SetForegroundColour( wx.Colour( 233, 0, 0 ) )
                self.m_staticText41.SetLabelText('请插入金税盘')
                self.usb_insert = False
            if first == True:
                self.reset_usb_status(self.usb_insert)
            elif self.his_usb_status != self.usb_insert:
                self.reset_usb_status(self.usb_insert)
        if self.usb_insert == False:
            self.remove_task()
            self.set_status('已准备')
        time.sleep(5)
        return self.check_usb_staus(False)

    # 重置 usb状态
    def reset_usb_status(self,usb_status):
        self.his_usb_status = self.usb_insert
        if self.usb_insert == True:
            self.app_status()
        ret = self.post_link(self.host + self.config['client']['update_usb_url'],{"invoice_export_client_id":self.uid,"usb_status":int(usb_status)})
        print(ret)

    # 修改客户端usb状态
    def update_usb_status(self):
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        try:
            res = requests.post(upload_link, data=data, files=files, headers=headers)
            return res.json()
        except:
            return {"code":500,"text":"网络异常，上传失败"}

    def connect_service(self):
        # self.run_status = bool(1 - self.run_status)
        if hasattr(self,'mythread') == False:
            self.err_connect = 0
            self.mythread = threading.Thread(target=self.create_websocket)
            self.cond = threading.Condition() # 锁
            self.mythread.start()

        # if self.run_status == True:
        #     # self.run_text.set('启动中')
        #     self.run_text.set('运行中')
        # else:
        #     delattr(self,'mythread')
        #     self.ws.close()
        #     self.run_text.set('停止中')

    def create_websocket(self):
        def on_message(ws, message):
            msg = json.loads(message)
            # msg = {"type":"action","room_id":"0da851c3bb31aaf458919479dcb726f0","send_to":"dd","data":"2019年 11月份 销项票数据导出"}
            if msg['type'] == 'action' and msg['data'] != '':
                # 默认消息弹框提醒
                self.tip = False
                self.m_listBox2.InsertItems([str(msg['data'])],0)
                self.running_index = 0
                self.parse_action(str(msg['data']))
                

        def on_error(ws, error):
            now = time.strftime('%S ',time.localtime(time.time()))
            if self.err_connect == 0:
                self.add_log('服务器连接已断开')
            self.err_connect += 1
            time.sleep(8)
            self.create_websocket()

        def on_close(ws):
            self.add_log('连接已断开')

        def on_open(ws):
            self.add_log('服务器连接成功')
            self.err_connect = 0
            uid,_ = regedit.get_pc_id()
            ws.send('{"type":"login","room_id":"%s","client_name":"%s"}' % (self.config['client']['room_id'],uid))

        websocket.enableTrace(True)
        ws = websocket.WebSocketApp("ws://im.itking.cc:12366",
                                    on_message = on_message,
                                    on_error = on_error,
                                    on_close = on_close)
        ws.on_open = on_open
        self.ws = ws
        ws.run_forever()

    # 检查用户代账服务状态
    def check_user_service(self,first):
        credit_code = self.config['log']['credit_code']
        check_url = self.host + self.config['client']['check_service']
        try:
            ret = requests.post(check_url,data={"credit_code":credit_code})
            check_user = ret.json()
        except:
            check_user = {"code":1,'text':'未能连接到服务端'}

        if check_user['code'] == 1:
            if first == True:
                dlg = wx.MessageDialog(None, check_user['text'] + '，是否继续运行', u"鑫山财务-开票辅助工具 提示", wx.YES_NO | wx.STAY_ON_TOP)
                if dlg.ShowModal() == wx.ID_NO:
                    return False
            else:
                dlg = wx.MessageDialog(None, check_user['text'], u"鑫山财务-开票辅助工具 提示", wx.OK | wx.STAY_ON_TOP)
                if dlg.ShowModal() == wx.ID_OK:
                    return False
        return True

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
        self.reset_by_data(ym)

    # 按指定时间设置界面
    def reset_by_data(self,ym):
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
        # 默认消息弹框提醒
        ym = self.get_y_m()
        
        if self.usb_insert == False:
            self.add_log('导出失败，检测到税控盘未插入')
            if self.tip == True:
                dlg = wx.MessageDialog(None, u"未检测到税控盘，请插入后重试", u"鑫山财务-开票辅助工具 提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
                if dlg.ShowModal() == wx.ID_OK:
                    pass
            return False

        auto_app = AutoExport.Export()
        # 最小化本应用窗口
        # auto_app.min_self()
        self.Hide()
        auto_app.do_ready()
        check_uer = self.check_user_service(False)
        if check_uer == False:
            return False
        # 修复数据
        
        try:
            auto_app.fix_data(ym)
            ret = auto_app.dw_excel(ym)
            auto_app.min_app()
        except:
            if self.tip == True:
                ret = {"code":500,"msg":"运行时发生了一个未知错误"}
            else:
                ret = {"code":500,"msg":"自动运行任务失败，请手动重试"}    
        # 导出上传数据
        self.add_log(ret['msg'])
        
        ym_split = ym.split('-')
        if ret['code'] == 200:
            self.add_log("进项票导出成功，正在上传")
            sales_name = self.config['upload']['sales_name'] + time.strftime("%Y%m%d",  time.localtime()) + self.config['upload']['sales_file_ext']
            sales_upload_link = self.host + self.config['upload']['sales_upload_link']
            post_data = {"credit_code":self.credit_code,"period":ym_split[0]+ym_split[1],"submit":"1","tax_import":"1"}
            ret = self.upload_export(sales_name,sales_upload_link,post_data,ym_split)
            if ret['code'] == 200:
                if type(self.running_index) == int:
                    self.m_listBox2.Delete(self.running_index)
                    self.running_index = False
                return ret
        elif ret['code'] == 404:
            self.add_exp_log({'content':'当前所属期 %s 没有进项票（航信）数据' % (ym_split[0]+ym_split[1]),"credit_code":self.credit_code,'action':'1','status':'1','period':ym_split[0]+ym_split[1]})
            if self.tip == True:
                dlg = wx.MessageDialog(None, ret['msg'], u"鑫山财务-开票辅助工具 提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
                if dlg.ShowModal() == wx.ID_OK:
                    self.m_listBox2.Delete(self.running_index)
                    pass
        else:
            if self.tip == True:
                dlg = wx.MessageDialog(None, ret['msg'] + " 是否重新运行", u"鑫山财务-开票辅助工具 发生错误", wx.YES_NO | wx.STAY_ON_TOP | wx.ICON_EXCLAMATION)
                if dlg.ShowModal() == wx.ID_YES:
                    return self.export_sales(event)
        self.add_exp_log({'content':'航信进项票数据(%s)导出失败' % (ym_split[0]+ym_split[1]),"credit_code":self.credit_code,'action':'1','status':'2','period':ym_split[0]+ym_split[1]})        
        self.running_index = False
        return {"code":500,"text":"上传失败"}


    def upload_export(self,sales_name,sales_upload_link,post_data,ym_split):
        up_res = self.upload_file(sales_name,sales_upload_link,post_data)
        if up_res['code'] == 0:
            self.add_exp_log({'content':'航信进项票数据(%s)导出并上传成功' % (ym_split[0]+ym_split[1]),"credit_code":self.credit_code,'action':'1','status':'1','period':ym_split[0]+ym_split[1]})
            if self.tip == True:
                dlg = wx.MessageDialog(None, u"上传成功", u"鑫山财务-开票辅助工具 提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
                if dlg.ShowModal() == wx.ID_OK:
                    # self.Close(True)
                    pass

            self.add_log('航信进项票(%s)上传成功' % (ym_split[0]+ym_split[1]))    
            return {"code":200,"text":"上传成功"}
        else:
            dlg = wx.MessageDialog(None, up_res['text'] + " 是否重新上传", u"鑫山财务-开票辅助工具 发生错误", wx.YES_NO | wx.STAY_ON_TOP | wx.ICON_EXCLAMATION)
            if dlg.ShowModal() == wx.ID_YES:
                return self.upload_export(sales_name,sales_upload_link,post_data,ym_split)
        return {"code":500,"text":"数据已导出，但上传失败"}     

  
    # 上传文件
    def upload_file(self,file_name,upload_link,data):
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        file_name = './exp_file/'+file_name
        if os.path.exists(file_name) == False:
            return {"code":500,"text":"未找到导出的数据文件"}
        files = {'upfile': open(file_name, 'rb')}
        try:
            res = requests.post(upload_link, data=data, files=files, headers=headers)
            return res.json()
        except:
            return {"code":500,"text":"网络异常，上传失败"}

    def toggle( self, event ):
        if self.m_radioBtn6.GetValue() == True:
            self.set_config('app','default',6)
        if self.m_radioBtn7.GetValue() == True:
            dlg = wx.MessageDialog(None, u"功能暂未支持，敬请期待", u"鑫山财务-开票辅助工具 提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
            if dlg.ShowModal() == wx.ID_OK:
                self.m_radioBtn6.SetValue(True)
                pass

    # 添加日志
    def add_log(self,text):
        if text == "":
            return False
        self.m_listBox3.InsertItems(["%s %s" % (time.strftime("%M:%S", time.localtime()),text)],0)

    # 添加任务
    def add_task(self):
        add_link = self.host + self.config['log']['task_url']
        task = self.post_link(add_link,{"credit_code":self.credit_code,'invoice_export_client_id':self.uid})

        if len(task) == 0 or 'code' in task:
            self.m_listBox2.InsertItems(["暂无任务"],0)
        else:
            self.m_listBox2.InsertItems(task,0)

    def remove_task(self):
        cc = self.m_listBox2.GetCount()
        for i in range(cc):
            self.m_listBox2.Delete(0)

    def post_link(self,link,post_data):
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        try:
            req = requests.post(link, data=post_data, headers=headers)
            resp = req.json()
            return resp
        except:
            return {"code":500,"text":"网络异常，请求失败"}


    # 开票软件运行状态
    def app_status(self):        
        self.set_config('log','last_run_time',time.strftime("%Y-%m-%d %H:%M", time.localtime()))
        # 自启动的 从配置文件获取用户信息
        print('检测启动方式')
        if '/autorun' in sys.argv:
            self.credit_code = self.config['log']['credit_code']
        else:
            self.credit_code = pname =''
            if self.config['app']['default'] == '6':
                print('打开航信信息客户端获取信息')
                auto_app = AutoExport.Export()
                auto_app.do_ready()
                self.credit_code,pname = auto_app.user_info()
        if pname != "":
            self.set_config('log','credit_code',self.credit_code)
            self.set_config('log','corpname',pname)
            self.set_status('开票软件已运行 | ' + pname)
            self.add_task()
        else:
            self.add_log("获取用户信息失败")
            if self.credit_code != '':
                credit_code = self.credit_code
            else:
                credit_code = self.config['log']['credit_code']
            if credit_code != '':
                print('登录失败,上传日志')
                self.add_exp_log({'content':'登录开票软件失败','credit_code':credit_code})

        # 如果从电脑获取的uid 跟本地uid不一致，则修改系统
        if self.config['client']['uid'] != self.uid and self.credit_code != '':
            self.set_config('client','uid',self.uid)

        # 检测上个月数据是否导出
        # today = datetime.date.today()
        # first = today.replace(day=1)
        # lastMonth = first - datetime.timedelta(days=1)
        # ym = lastMonth.strftime("%Y%m")

        # export = self.config['export']['sales_list']
        # export_split = export.split(",")
        # if ym in export_split:
        #     self.m_staticText5.SetLabelText('上月数据已导出')
        #     self.m_staticText5.SetForegroundColour( wx.Colour( 76, 146, 76 ) )
        # else:    
        #     self.m_staticText5.SetLabelText('上月数据还未上传')
        
        

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

    # 任务框点击处理
    def task_click(self,event):
        # 正在运行
        listbox = event.GetEventObject()
        select_list = listbox.GetStringSelection()
        self.running_index = listbox.GetSelection()
        # 默认消息弹框提醒
        self.tip = True
        self.parse_action(select_list)
              

    def parse_action(self,actin_text):
        sel_split = actin_text.split(' ')
        if len(sel_split) < 2:
            return False

        y = sel_split[0].replace('年','')
        m = sel_split[1].replace('月份','')
        m = '%02d' % int(m)
        self.reset_by_data(y+'-'+m)
        if '销项' in sel_split[2]:
            if '航天' in sel_split[2]:
                return self.export_sales(None)
            else:
                dlg = wx.MessageDialog(None, u"功能暂未支持，敬请期待", u"鑫山财务-开票辅助工具 提示", wx.OK | wx.STAY_ON_TOP | wx.ICON_INFORMATION)
                if dlg.ShowModal() == wx.ID_OK:
                    pass   

    def add_exp_log(self,data):
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        add_link = self.host + self.config['log']['post_url']
        return requests.post(add_link, data=data, headers=headers)

app = wx.App()
frame = Ep(None)
ready = frame.start()
if ready == False:
    sys.exit()

if '/autorun' in sys.argv:
    show_fram = False
else:
    show_fram = True

frame.Show(show_fram)
app.MainLoop()