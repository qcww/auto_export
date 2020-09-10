#-*- coding: utf-8 -*-

from pywinauto.application import Application
import pywinauto
import win32clipboard as w
import random
import time
import datetime
import json
import os
import re
import sys
import winreg
import configparser
import win32api
import win32con
import math
import locale
import regedit
import requests

class Export:
    def __init__(self):
        # config_path = os.path.abspath(os.path.dirname(__file__))
        # os.chdir(config_path)
        self.config_file = "./config/config.ini"
        if os.path.exists(self.config_file) == False:
            self.config_file = regedit.get_client_path()+"\\config\\config.ini"

        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')
        locale.setlocale(locale.LC_ALL,'en')
        locale.setlocale(locale.LC_CTYPE,'chinese')
        self.menu_action = ['系统设置','发票管理','汇总处理','系统维护']
        self.export_menu = ["汇总管理->发票数据导出->发票数据导出","汇总管理->发票数据导出->清单发票数据导出"]
        self.do_ready()
        # self.user_info()
        print('检测版本兼容信息')

    def post_link(self,link,post_data):
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        try:
            req = requests.post(link, data=post_data, headers=headers)
            resp = req.json()
            return resp
        except:
            return {"code":500,"text":"网络异常，请求失败"}
        
        
    def run_app(self,timeout):
        try:
            print('重复尝试检测登录窗口')
            self.app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app")
            ac = self.app.window(title="税控发票开票软件")
            ac.wait('exists', timeout=15, retry_interval=1)
            ac.set_focus()
        except:
            timeout -= 1
            if timeout > 0:
                print('重试')
                time.sleep(10)
                return self.run_app(timeout)

    # 最小化本应用窗口
    def min_app(self):
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
        try:
            ac.minimize()
        except:
            pass

    def login(self):
        # 已经打开
        try:
            self.app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            return True
        except:
            print('开票未打开')
            pass
        # 打开了登录框
        try:
            self.app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id="LoginForm")
            ac = self.app.windows()[0]
            ac.set_focus()
        except:
            print('未检测到登录窗口,需要启动软件')
            ht_app = self.get_app_path()
            if ht_app == '':
                return False
            app = Application(backend='uia').start(ht_app,timeout=15,retry_interval=5)
            time.sleep(10)
            print('等待程序启动')
            self.run_app(3)
        # 需要点击登录
        print('尝试登录')
        self.try_login(True)
        return self.tax_qk()

    def get_qk_status(self):
        print('检查清卡时间')
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id="MDIMainForm")
        ztcx_btn = ac.child_window(auto_id="btnZTCX", control_type="Button")
        if ztcx_btn.exists() == False:
            select_menu = 2
            self.tab_menu(select_menu)
            time.sleep(1)
        
        if ztcx_btn.exists() == False:
            return False

        ztcx_btn.set_focus()    
        try:
            ztcx_btn.click()
        except:
            pass
        
        tab_btn = ac.child_window(title="增值税专用发票及增值税普通发票", control_type="TabItem")
        tab_btn.wait('exists', timeout=2, retry_interval=5)
        tab_btn.select()
        ss_time = ac.child_window(auto_id="lblLockday",class_name_re="WindowsForms10.STATIC")
        ss_time.wait('exists', timeout=2, retry_interval=5)
        s_end_time = ss_time.window_text()
        print('锁死时间',s_end_time)
        pywinauto.keyboard.send_keys('%{F4}')
        return s_end_time

    # 默认锁死日期
    def get_next_month(self,year=None,month=None):
        # 默认本年本月
        if(year == None or month == None):
            today = datetime.datetime.today()
            year = today.year
            month = today.month
        if month == 12:
            month = 1
            year += 1
        else:
            month += 1
        return datetime.datetime(year,month,1).strftime("%Y%m")

    # 上传清卡状态
    def tax_qk(self):
        time.sleep(3)
        if self.app['系统参数设置'].exists() == True:
            print('检测到需要初始设置')
            return False

        print('检测到需要汇总上传')    
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",title="汇总信息")
        if ac.exists() == True:
            try:
                ac['确认'].click()
            except:
                pass
        else:
            return True
        
        time.sleep(5)
        
        mes_win = self.app.SysMessageBox.child_window(auto_id="lblMsg", control_type="Text")
        try:
            print('等待清卡完成')
            mes_win.wait_not('exists', timeout=20, retry_interval=5)
        except:
            print('等待清卡任务完成超时了')
            return False
        
        # 关闭其它弹框
        if self.app.SysMessageBox.exists() == True:
            # 完成清卡操作
            mes = self.app.SysMessageBox.child_window(auto_id="lblMsg", control_type="Text")
            print('检测到清卡完成信息')
            if '完成清卡' in mes.texts()[0]:
                try:
                    self.post_link(self.config['link']['host'] + self.config['upload']['tax_qk'],{"period":self.get_next_month(),"credit_code":"%s" % self.credit_code})
                except :
                    pass
            if self.app.SysMessageBox.exists == True:    
                self.app.SysMessageBox['确认'].click()
                print('点击了确认信息框，稍等')
            time.sleep(5)
        return True

    # 登录检测
    def try_login(self,retry):
        print('需要登录信息')
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id="LoginForm")
        if ac.window(title="登录", auto_id="btnOK", control_type="Button").exists() == True:
            print('检测到登录按钮')
            try:
                title_win = ac.window(class_name_re="WindowsForms10.Window.8.app",auto_id="BodyClient")
                credit_text = title_win.window_text()
                credit_code = credit_text.split(" ")[-1]
                uid,_ = regedit.get_pc_id()
                if credit_code != '' and uid != '':
                    print('绑定电脑id')
                    self.post_link(self.config['link']['host'] + self.config['client']['update_client_id_url'],{"credit_code":"%s" % credit_code,"invoice_client_id":"%s" % uid})        
            except:
                print('设备id绑定失败')

            if retry == False:
                if ac.CheckBox.get_toggle_state() == 0:
                    # 未知原因导致报错，需要屏蔽
                    try:
                        ac.CheckBox.click()
                    except:
                        pass

                    ac.SysMessageBox.wait('exists', timeout=5, retry_interval=1)
                    ac.SysMessageBox.window(title="确认", auto_id="btnYes", control_type="Button").click()
                try:
                    title_win = ac.window(class_name_re="WindowsForms10.Window.8.app",auto_id="BodyClient")
                    credit_text = title_win.window_text()
                    self.credit_code = credit_text.split(" ")[-1]
                    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
                    req = requests.post(self.config['link']['host'] + self.config['app']['jsp_pwd_url'],{"credit_code":self.credit_code}, headers=headers)
                    req_json = req.json()
                    print('从服务端获取的密码口令',req_json)
                    pwd = req_json['tax_disk_pwd'] if req_json['tax_disk_pwd'] != '' else self.config['client']['pwd']
                    cert = req_json['tax_disk_shibboleth'] if req_json['tax_disk_shibboleth'] != '' else self.config['client']['cert']
                except:
                    pwd = self.config['client']['pwd']
                    cert = self.config['client']['cert']
                print("默认账号密码",pwd,cert)
                self.set_config('client','pwd',pwd)
                self.set_config('client','cert',cert)
                # 尝试从网站获取配置
                pwd_win = ac.window(auto_id="txtPwd", control_type="Edit")
                time.sleep(2)
                pwd_win.set_text(pwd)
                cert_win = ac.window(auto_id="txtCertPassword", control_type="Edit")
                cert_win.set_text(cert)
            # 未知原因需要屏蔽错误 
            print('点击登录按钮')
            try:
                ac["登录"].click() 
            except:
                pass    
              
        time.sleep(5)
        if ac.window(title="登录错误", auto_id="BodyBounds", control_type="Pane").exists() == True and retry == True:
            # 重试
            print('重试一次')
            err_win = ac.window(title="登录错误", auto_id="BodyBounds", control_type="Pane")
            err_button = err_win.child_window(title="确认", auto_id="btnYes", control_type="Button")
            err_button.wait('exists', timeout=5, retry_interval=1)
            err_button.click()
            if retry == True:
                return self.try_login(False)

    def get_app_path(self):
        app_path = self.config['app']['path_ht']
        if app_path == "":
            try:
                reg_key = winreg .OpenKey(winreg .HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\fwkp.exe")
                app_path_all, _ = winreg.QueryValueEx(reg_key, "Path")

                path_start = app_path_all.find('开票软件')
                app_path = app_path_all[0:int(path_start)+5] + self.config['app']['name_ht']
            except:
                app_path = ''
        else:
            app_path = self.config['app']['path_ht'] + self.config['app']['name_ht']
        if app_path != '' and os.path.exists(app_path):
            return app_path
        return ''

    # 写入配置
    def set_config(self,section,key,value):
        set_val = str(value)
        self.config.set(section,key,re.sub(r'%','#53',set_val.strip('|')))
        self.config.write(open(self.config_file,'w',encoding='utf-8'))

    def user_info(self):
        print('获取登录用户的客户端信息')
        try:
            ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            status_bar = ac.window(auto_id="statusStrip1", control_type="StatusBar")
            status_bar.wait('exists', timeout=10, retry_interval=1)
            match_uid = status_bar.children()[0].texts()[0]
            p1 = match_uid.find("#")
            p2 = match_uid.find(".")
            uid = match_uid[int(p1)+1:int(p2)]
            corpname = status_bar.children()[1].texts()[0]
            # ac.minimize()
            return uid,corpname
        except:    
            return '',''

    # 版本兼容处理（仅仅处理已知版本兼容问题）
    def check_version(self):
        print('兼容处理部分代码')
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
        # v2.2.34版本出现
        if ac.Toolbar.window(title="报税处理").exists() == True:
            self.menu_action[2] = "报税处理"
            self.export_menu[0] = "报税处理->发票数据导出->发票数据导出"

    def tab_menu(self,tab_menu):
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
        # dlg.print_control_identifiers()
        tool_bar = ac.window(auto_id="toolStripMenu", control_type="ToolBar")
        hz_button = tool_bar.window(title=self.menu_action[tab_menu], control_type="Button")
        hz_button.click()
        time.sleep(3)

    # 计算月份差 
    def months(self,date1,date2):
        date1_splite = date1.split('-')
        date2_splite = date2.split('-')
        y = int(date2_splite[0]) - int(date1_splite[0])
        m = int(date2_splite[1]) - int(date1_splite[1])
        return y*12 + m

    def min_self(self):
        try:
            app = Application(backend='uia').connect(class_name_re="wxWindowNR",title_re = "鑫山财务")
            ac = app.windows()[0]
            ac.minimize()
        except:
            pass


    # 执行准备动作
    def do_ready(self):
        print('do ready')
        # 防止不规范操作导致的关闭应用或最小化窗口
        self.login()

        # 关闭重复弹框的提醒框
        try:
            print('检测重复运行弹框，并关闭')
            run_over = Application(backend='uia').connect(title="CusMessageBox")
            if run_over.window(title="CusMessageBox").exists() == True:
                run_over.window(title="CusMessageBox").close()
        except:
            pass

        try:
            app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id="MDIMainForm")
            ac = app.windows()
            # 关闭其它弹框
            print('关闭多余窗口')
            if len(ac) > 1:
                for i in ac:
                    if i.automation_id() != 'MDIMainForm':
                        i.close()
            ac = app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            if ac.SysMessageBox.exists() == True:
                ac.SysMessageBox.close()
            ac.maximize()
            ac.set_focus()
        except:
            pass

    def getClipboardText(self):
        w.OpenClipboard()
        d = w.GetClipboardData(win32con.CF_TEXT)
        w.CloseClipboard()
        return(d).decode('GBK')

    # 修复数据
    def fix_data(self,ym):
        # 切换菜单
        select_menu = 1
        self.tab_menu(select_menu)
        try:
            app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title_re="增值税发票税控开票软件",auto_id = "MDIMainForm")
            ac = app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            ac.menu_select(u"发票管理->发票修复")
        except:
            pass

        fix_win = ac.window(title="发票修复", auto_id="SelectMonth")
        fix_win.wait('exists', timeout=10, retry_interval=1)

        ym_split = ym.split('-')
        now = time.strftime("%Y-%m-01",  time.localtime())
        now_split = now.split('-')
        dec_year = int(now_split[0]) - int(ym_split[0])
        if dec_year > 0:
            dec_month = 12 - int(ym_split[1])
        else:
            dec_month = int(now_split[1]) - int(ym_split[1])

        fix_win.window(class_name_re="WindowsForms10.COMBOBOX.app.",auto_id="aisinoCMB_Year").set_focus()
        if dec_year > 0:
            for i in range(dec_year):
                win32api.keybd_event(40,0,0,0)
                win32api.keybd_event(40,0,win32con.KEYEVENTF_KEYUP,0)

        fix_win.window(auto_id="com_month", control_type="ComboBox").set_focus()
        if dec_month > 0:
            for i in range(dec_month):
                win32api.keybd_event(40,0,0,0)
                win32api.keybd_event(40,0,win32con.KEYEVENTF_KEYUP,0)

        fix_win.window(title="确定", auto_id="but_ok", control_type="Button").click()

        time.sleep(2)
        fix_win = app.window(title="发票修复过程")
        fix_win.wait_not('exists', timeout=60, retry_interval=3)
        time.sleep(1)

        app = Application(backend='uia').connect(title="SysMessageBox",class_name_re="WindowsForms10.Window.8.app")
        mes_win = app.window(title="SysMessageBox")
        mes_win.wait('exists', timeout=10, retry_interval=1)
        mes_win.window(title="确认", class_name_re="WindowsForms10.BUTTON.app").click()

    # 导出excel
    def dw_excel(self,ym,index):
        select_menu = 2
        self.tab_menu(select_menu)
        time.sleep(1)
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
        try:
            ac.maximize()
            ac.set_focus()
            ac.menu_select(self.export_menu[index])
            print(self.export_menu[index])
        except:
            pass
        
        now = time.strftime("%Y-%m-01",  time.localtime())
        m = self.months(ym,now)
        time.sleep(1)
        try:
            ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
            mon_win = ac.window(class_name_re="WindowsForms10.Window.8.app",auto_id="FPExport")
            # mon_win.print_control_identifiers()
            mon_sel = mon_win.window(title="月份", auto_id="cmbMonth")
            mon_sel.set_focus()

            if m > 0:
                for i in range(m):
                    win32api.keybd_event(40,0,0,0)     # enter
                    win32api.keybd_event(40,0,win32con.KEYEVENTF_KEYUP,0)
            time.sleep(1)
            # 检查导出的数据日期是否正确
            ym_sp = ym.split('-')
            if mon_win.window(title="%d年%d月%d日" % (int(ym_sp[0]),int(ym_sp[1]),int(ym_sp[2])), auto_id="dtpStart").exists() == True:
                ok_btn = mon_win.child_window(title="确定",found_index=0, auto_id="btnOK", control_type="Button")
                if ok_btn.exists() == True:
                    try:
                        ok_btn.click()
                    except:
                        pass
                    time.sleep(1)
            else:
                return {"code":301,"msg":"超出所选期或无数据"}
        except:
            pass

        try:
            app = Application(backend='uia').connect(title="SysMessageBox",class_name_re="WindowsForms10.Window.8.app")
            if app.window(title="SysMessageBox").exists() == True:
                mes_win = app.window(title="SysMessageBox")
                mes_win.wait('exists', timeout=10, retry_interval=1)
                if mes_win.window(title="您选择的时间段没有数据", auto_id="lblMsg").exists() == True:
                    mes_win.window(title="确认", auto_id="btnYes").click()
                    return {"code":404,"msg":"您选择的时间段没有数据"}
        except:
            pass

        try:
            # 另存
            save_win = self.app.window(title='另存为')
            save_win.wait('exists', timeout=5, retry_interval=1)

            pywinauto.keyboard.send_keys('^n^x')
            upload_file_name = self.getClipboardText()
            fix_name = "%d年%d月-%d" % (int(ym_sp[0]),int(ym_sp[1]),random.randint(1000,9999)) + upload_file_name
            if 'xlsx' not in fix_name:
                fix_name += '.xlsx'

            pwd = regedit.get_client_path()+"\\exp_file\\" + fix_name
            print('导出的文件名',pwd)
            save_win.ComboBox.Edit.set_text(pwd)
            pywinauto.keyboard.send_keys("^n%s")

            if save_win.window(title="确认另存为").exists() == True:
                save_win.window(title="确认另存为").child_window(title_re="是").click()

            time.sleep(2)

            ac = self.app.window(auto_id="FPProgressBar",class_name_re="WindowsForms10.Window.8.app")
            ac.wait('exists', timeout=5, retry_interval=1)
            ac.set_focus()
            mes_win = ac.window(title="SysMessageBox")
            mes_win.wait('exists', timeout=10, retry_interval=1)
            if ac.window(title="导出成功").exists() == True:
                ac.window(title="确认").click()
                return {"code":200,"msg":"","upload_file_name":fix_name}
        except:
            pass


        return {"code":500,"msg":""}