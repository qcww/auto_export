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
        self.login()
        self.user_info()
        
    def run_app(self,timeout):
        try:
            time.sleep(5)
            app = Application(backend='uia').connect(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1",title="税控发票开票软件")
        except:
            timeout -= 1
            return self.run_app(timeout)
        return app

    def login(self):
        try:
            app = Application(backend='uia').connect(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1")
        except:
            ht_app = self.config['app']['path_ht']
            app = Application(backend='uia').start(ht_app)
            app = self.run_app(5)
        # print(app.windows())
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

    def user_info(self):
        # select_menu = 0
        # self.tab_menu(select_menu)
        ac = self.app.window(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1")
        status_bar = ac.window(auto_id="statusStrip1", control_type="StatusBar")
        match_uid = status_bar.children()[0].texts()[0]
        p1 = match_uid.find("#")
        p2 = match_uid.find(".")
        uid = match_uid[int(p1)+1:int(p2)]
        print(uid)

    def tab_menu(self,tab_menu):
        ac = self.app.window(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1")
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
        ac = self.app.window(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1")
        ac.menu_select("汇总管理->发票数据导出->发票数据导出")

# menu =ac.child_window(auto_id="toolStripMenu", control_type="System.Windows.Forms.ToolStrip")
# menu.draw_outline(colour ='green',thickness = 2,rect = None)

auto = AutoExport()
auto.dw_excel()