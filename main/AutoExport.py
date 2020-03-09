#-*- coding: utf-8 -*-

from pywinauto.application import Application
import pywinauto
import time
import datetime
import json
import os
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
            self.config_file = regedit.get_client_path()+"\\main\\config\\config.ini"

        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')
        locale.setlocale(locale.LC_ALL,'en') 
        locale.setlocale(locale.LC_CTYPE,'chinese')
        self.menu_action = ['系统设置','发票管理','汇总处理','系统维护']
        self.export_menu = "汇总管理->发票数据导出->发票数据导出"
        self.do_ready()
        # self.user_info()
        print('检测版本兼容信息')
        self.check_version()
        
    def run_app(self,timeout):
        # try:
        #     print('检测已登录窗口')
        #     app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
        #     ac = app.window(found_index=0)
        #     ac.set_focus()
        #     return app
        # except:
        #     print('未检测到已登录窗口，需要登录')
        #     pass 

        try:
            print('重复尝试检测登录窗口')
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
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
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
            pass    
        # 打开了登录框
        try:
            print('检测登录窗口')
            self.app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title="税控发票开票软件")
            ac = app.windows()[0]
            ac.set_focus()
        except:
            ht_app = self.get_app_path()
            if ht_app == '':
                return False
            app = Application(backend='uia').start(ht_app,timeout=15,retry_interval=5)
            time.sleep(10)
            print('等待程序启动')
            self.app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title="税控发票开票软件")
            self.run_app(3)
            # 需要点击登录
            self.try_login(True)
        # try:
        #     self.try_login(True)
        # except:
        #     ac = app.windows()[0]
        #     ac.set_focus()
        # self.app = app

    # 登录检测
    def try_login(self,retry):
        print('需要登录信息')
        ac = self.app.window(title="税控发票开票软件")
        if ac.window(title="登录", auto_id="btnOK", control_type="Button").exists() == True:
            if retry == False:
                if ac.CheckBox.get_toggle_state() == 0:
                    # 未知原因导致报错，需要屏蔽
                    try:
                        ac.CheckBox.click()
                    except:
                        pass
                    
                    print('安全提示')
                    ac.SysMessageBox.wait('exists', timeout=5, retry_interval=1)
                    print('关闭安全提示')
                    ac.SysMessageBox.window(title="确认", auto_id="btnYes", control_type="Button").click()
                try:
                    title_win = ac.window(class_name_re="WindowsForms10.Window.8.app",auto_id="BodyClient")
                    credit_text = title_win.window_text()
                    credit_code = credit_text.split(" ")[-1]
                    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
                    req = requests.post(self.config['link']['host'] + self.config['app']['jsp_pwd_url'],{"credit_code":credit_code}, headers=headers)
                    req_json = req.json()
                    pwd = req_json['tax_disk_pwd']
                    cert = req_json['tax_disk_shibboleth']
                except:
                    pwd = '123456'
                    cert = '12345678'

                # 尝试从网站获取配置
                pwd = ac.window(auto_id="txtPwd", control_type="Edit")
                time.sleep(2)
                text_pw = pwd.texts()
                pwd.set_text("123456")
                cert = ac.window(auto_id="txtCertPassword", control_type="Edit")
                cert.set_text('12345678')
            # 未知原因需要屏蔽错误 
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

    def user_info(self):
        # select_menu = 0
        # self.tab_menu(select_menu)
        try:
            ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            status_bar = ac.window(auto_id="statusStrip1", control_type="StatusBar")
            status_bar.wait('exists', timeout=10, retry_interval=1)
            match_uid = status_bar.children()[0].texts()[0]
            p1 = match_uid.find("#")
            p2 = match_uid.find(".")
            uid = match_uid[int(p1)+1:int(p2)]
            corpname = status_bar.children()[1].texts()[0]
            ac.minimize()
            return uid,corpname
        except:    
            return '',''

    # 版本兼容处理（仅仅处理已知版本兼容问题）
    def check_version(self):
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
        # v2.2.34版本出现
        if ac.Toolbar.window(title="报税处理").exists() == True:
            self.menu_action[2] = "报税处理"
            self.export_menu = "报税处理->发票数据导出->发票数据导出"

    def tab_menu(self,tab_menu):
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
        # dlg.print_control_identifiers()
        tool_bar = ac.window(auto_id="toolStripMenu", control_type="ToolBar")
        hz_button = tool_bar.window(title=self.menu_action[tab_menu], control_type="Button")
        hz_button.click()
        time.sleep(3)

    # 计算月份差 
    def months(self,date1,date2):
        date1=time.strptime(date1,"%Y-%m-%d")
        date2=time.strptime(date2,"%Y-%m-%d")
        date1=datetime.datetime(date1[0],date1[1],date1[2])
        date2=datetime.datetime(date2[0],date2[1],date2[2])
        #返回两个变量相差的值，就是相差天数
        return math.floor((date2-date1).days/30)

    def min_self(self):
        try:
            app = Application(backend='uia').connect(class_name_re="wxWindowNR",title_re = "鑫山财务")
            ac = app.windows()[0]
            ac.minimize()
        except:
            pass


    # 执行准备动作
    def do_ready(self):
        # 防止不规范操作导致的关闭应用或最小化窗口
        try:
            self.login()
        except:
            pass
        
        try:
            app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title_re="增值税发票税控开票软件",auto_id = "MDIMainForm")
            ac = app.windows()
            # 关闭其它弹框
            if len(ac) > 1:
                for i in ac:
                    if i.automation_id() != 'MDIMainForm':
                        i.close()

            ac = app.windows()[0]
            ac.maximize()
            ac.set_focus()
        except:
            pass


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
    def dw_excel(self,ym):
        select_menu = 2
        self.tab_menu(select_menu)
        time.sleep(1)
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
        try:
            ac.maximize()
            ac.set_focus()
            ac.menu_select(self.export_menu)
        except:
            pass
        
        now = time.strftime("%Y-%m-01",  time.localtime())
        m = self.months(ym,now)
        try:
            app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            ac = app.window(class_name_re="WindowsForms10.Window.8.app")
            mon_win = ac.window(title="发票数据导出",auto_id="FPExport")
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
                export_window = ac.window(title="发票数据导出", auto_id="FPExport", control_type="Window")
                export_window.window(title="确定", auto_id="btnOK").click()
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
            app = Application(backend='uia').connect(title="另存为")
            save_win = app.window(title='另存为')
            save_win.wait('exists', timeout=10, retry_interval=1)
            save_win.set_focus()
            time.sleep(1)
            save_win.child_window(title="上一个位置", control_type="Button").click()
            pwd = regedit.get_client_path()+'\main\exp_file'
            # print(pwd)
            save_win.child_window(title="地址", class_name="Edit").set_text(pwd)

            pywinauto.keyboard.send_keys("{ENTER 1}")

            save_win.window(title_re="保存",class_name="Button").click()
            if save_win.window(title="确认另存为").exists() == True:
                save_win.window(title="确认另存为").child_window(title_re="是").click()

            time.sleep(2)
            app = Application(backend='uia').connect(title="发票数据导出")
            ac = app.window(title="发票数据导出")
            ac.wait('exists', timeout=5, retry_interval=1)
            ac.set_focus()
            mes_win = ac.window(title="SysMessageBox")
            mes_win.wait('exists', timeout=10, retry_interval=1)
            if ac.window(title="导出成功").exists() == True:
                ac.window(title="确认").click()
                return {"code":200,"msg":""}
        except:
            pass


        return {"code":500,"msg":""}