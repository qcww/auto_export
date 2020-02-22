#-*- coding: utf-8 -*-

from pywinauto.application import Application
import pywinauto
import time
import json
import configparser

class AutoExport:
    def __init__(self):
        self.config_file = './config/config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')
        self.menu_action = ['系统设置','发票管理','汇总处理','系统维护']
        self.export_menu = "汇总管理->发票数据导出->发票数据导出"
        self.login()
        self.user_info()
        self.check_version()
        
    def run_app(self,timeout):
        try:
            time.sleep(5)
            app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title="税控发票开票软件")
        except:
            timeout -= 1
            return self.run_app(timeout)
        return app

    def login(self):
        # 已经打开
        try:
            self.app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
            return True
        except:
            pass    
        # 打开了登录框
        try:
            app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title="税控发票开票软件")
        except:
            ht_app = self.get_app_path()
            app = Application(backend='uia').start(ht_app)
            app = self.run_app(5)
        # print(app.windows())
        # 需要点击登录
        try:
            ac = app.window(title="税控发票开票软件")
            # ac.print_control_identifiers()
            ac.set_focus()
            ac["登录"].click()
            time.sleep(2)
        except:
            ac = app.windows()[0]
            ac.set_focus()
        self.app = app

    def get_app_path(self):
        self.config['app']['path_ht']    

    def user_info(self):
        # select_menu = 0
        # self.tab_menu(select_menu)
        try:
            ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
            status_bar = ac.window(auto_id="statusStrip1", control_type="StatusBar")
            match_uid = status_bar.children()[0].texts()[0]
            p1 = match_uid.find("#")
            p2 = match_uid.find(".")
            uid = match_uid[int(p1)+1:int(p2)]
            print(uid)
        except:    
            return False

    # 版本兼容处理（仅仅处理已知版本兼容问题）
    def check_version(self):
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
        # v2.2.34版本出现
        if ac.Toolbar.window(title="报税处理").exists() == True:
            self.menu_action[2] = "报税处理"
            self.export_menu = "报税处理->发票数据导出->发票数据导出"
            print('Yes')

    def tab_menu(self,tab_menu):
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
        ac.set_focus()
        ac.maximize()
        # dlg.print_control_identifiers()
        tool_bar = ac.window(auto_id="toolStripMenu", control_type="ToolBar")

        hz_button = tool_bar.window(title=self.menu_action[tab_menu], control_type="Button")
        hz_button.click()
        time.sleep(1)

    def dw_excel(self):
        select_menu = 2
        self.tab_menu(select_menu)
        ac = self.app.window(class_name_re="WindowsForms10.Window.8.app")
        ac.menu_select(self.export_menu)

# menu =ac.child_window(auto_id="toolStripMenu", control_type="System.Windows.Forms.ToolStrip")
# menu.draw_outline(colour ='green',thickness = 2,rect = None)

auto = AutoExport()
auto.dw_excel()