from pywinauto.application import Application
import pywinauto
import win32clipboard as w
import time
import json
import configparser
import win32api
import win32con
import datetime
import math
import locale
import os
import requests
import wx
import regedit

# app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app")
# ac = app.window(class_name_re="WindowsForms10.Window.8.app")
# mon_win = ac.window(title="发票数据导出",auto_id="FPExport")
# # mon_win.print_control_identifiers()
# mon_sel = mon_win.window(title="月份", auto_id="cmbMonth")
# mon_sel.set_focus()
# mon_sel.set_focus()

# for i in range(3):
#     win32api.keybd_event(40,0,0,0)     # enter
#     win32api.keybd_event(40,0,win32con.KEYEVENTF_KEYUP,0)

# app = Application(backend='uia').connect(class_name="WindowsForms10.Window.8.app.0.15a303f_r12_ad1",title="税控发票开票软件")
# ac = app.windows()[0]
# ac.set_focus()

# app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
# ac = app.window(class_name_re="WindowsForms10.Window.8.app")
# ac.set_focus()
# export_window = ac.window(title="发票数据导出", auto_id="FPExport", control_type="Window")
# export_window.window(title="确定", auto_id="btnOK").click()


# app = Application(backend='uia').connect(title="SysMessageBox",class_name_re="WindowsForms10.Window.8.app")

# if app.window(title="SysMessageBox").exists() == True:
#     mes_win = app.window(title="SysMessageBox")
#     if mes_win.window(title="您选择的时间段没有数据", auto_id="lblMsg").exists() == True:
#         print('没有数据')
#         mes_win.window(title="确认", auto_id="btnYes").click()
# time.sleep(1)
# print('1')




# app = Application(backend='uia').connect(title="另存为")
# save_win = app.window(title='另存为')
# save_win.wait('exists', timeout=30, retry_interval=1)
# save_win.set_focus()
# time.sleep(1)
# save_win.child_window(title="上一个位置", control_type="Button").click()
# pwd = os.getcwd()+'\exp_file'
# # print(pwd)
# save_win.child_window(title="地址", class_name="Edit").set_text(pwd)

# pywinauto.keyboard.send_keys("{ENTER 1}")
# time.sleep(3)
# pywinauto.keyboard.send_keys("{ENTER 1}")
# save_win.window(title_re="保存",class_name="Button").click()
# if save_win.window(title="确认另存为").exists() == True:
#     save_win.window(title="确认另存为").child_window(title_re="是").click()

# 监测导出成功
# app = Application(backend='uia').connect(title="发票数据导出")
# ac = app.window(title="发票数据导出")
# mes_win = ac.window(title="SysMessageBox")
# mes_win.wait('exists', timeout=10, retry_interval=1)
# if ac.window(title="导出成功").exists() == True:
#     ac.window(title="确认").click()



# ym = "2019-12-01".split('-')

# app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
# ac = app.window(class_name_re="WindowsForms10.Window.8.app")
# mon_win = ac.window(title="发票数据导出",auto_id="FPExport")
# if mon_win.window(title="%d年%d月%d日" % (int(ym[0]),int(ym[1]),int(ym[2])), auto_id="dtpStart").exists() == True:
#     print('yes')

# 发票修复
# try:
#     app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app.",auto_id = "MDIMainForm")
#     ac = app.window(class_name_re="WindowsForms10.Window.8.app.")
#     ac.menu_select(u"发票管理->发票修复")
# except:
#     pass

# fix_win = ac.window(title="发票修复", auto_id="SelectMonth")
# fix_win.wait('exists', timeout=10, retry_interval=1)

# ym_split = "2019-05-01".split('-')
# now = time.strftime("%Y-%m-01",  time.localtime())
# now_split = now.split('-')
# dec_year = int(now_split[0]) - int(ym_split[0])
# if dec_year > 0:
#     dec_month = 12 - int(ym_split[1])
# else:
#     dec_month = int(now_split[1]) - ym_split[1]

# fix_win.window(class_name_re="WindowsForms10.COMBOBOX.app.",auto_id="aisinoCMB_Year").set_focus()
# if dec_year > 0:
#     for i in range(dec_year):
#         win32api.keybd_event(40,0,0,0)
#         win32api.keybd_event(40,0,win32con.KEYEVENTF_KEYUP,0)

# fix_win.window(auto_id="com_month", control_type="ComboBox").set_focus()
# if dec_month > 0:
#     for i in range(dec_month):
#         win32api.keybd_event(40,0,0,0)
#         win32api.keybd_event(40,0,win32con.KEYEVENTF_KEYUP,0)

# fix_win.window(title="确定", auto_id="but_ok", control_type="Button").click()

# time.sleep(2)
# fix_win = app.window(title="发票修复过程")
# fix_win.wait_not('exists', timeout=60, retry_interval=3)
# time.sleep(1)

# app = Application(backend='uia').connect(title="SysMessageBox",class_name_re="WindowsForms10.Window.8.app")
# mes_win = app.window(title="SysMessageBox")
# mes_win.wait('exists', timeout=10, retry_interval=1)
# mes_win.window(title="确认", class_name_re="WindowsForms10.BUTTON.app").click()


# 上传文件
# def upload_file(file_name,upload_link):
#     headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
#     file_name = './exp_file/'+file_name
#     if os.path.exists(file_name):
#         print('fff')
#     files = {'upfile': open(file_name, 'rb')}
#     data = {'credit_code': '91340100MA2NPN203B', 'period': '202003', 'submit': '1', 'cate': '1'}
#     res = requests.post(upload_link, data=data, files=files, headers=headers)
#     return res

# re = upload_file('增值税专普清单发票数据导出20200402.xlsx','http://tinterface.hfxscw.com/interface.php?r=tax/upload-account-sales')
# print(re.content)


# config_file = './config/config.ini'
# config = configparser.ConfigParser()
# config.read(config_file,encoding='utf-8')

# export = config['export']['sales_list']
# if export != "":
#     sales_list = export.split(",")
#     sales_list.append('fde')
#     print(','.join(map(str,sales_list)))
# else:
#     print('fff')


# def add_exp_log(data):
#     config_file = './config/config.ini'
#     config = configparser.ConfigParser()
#     config.read(config_file,encoding='utf-8')

#     headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
#     add_link = config['log']['post_url']
#     return requests.post(add_link, data=data, headers=headers)


# add_res = add_exp_log({'content':'2019年 11月份 销项票数据导出','credit_code':'913401003551536121','source':'1','corpid':'83','action':'2','period':'201911'})
# print(add_res.content)

# 登录检测
# def try_login(retry):
#     app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title="税控发票开票软件")
#     ac = app.window(found_index=0)
#     if app.window(found_index=0).child_window(title="登录", auto_id="btnOK", control_type="Button").exists() == True:
        
#         ac.set_focus()
        
#         if retry == False:
#             pwd = ac.window(auto_id="txtPwd", control_type="Edit")
#             time.sleep(2)
#             text_pw = pwd.texts()
#             pwd.set_text("123456")
#             cert = ac.window(auto_id="txtCertPassword", control_type="Edit")
#             cert.set_text('12345678')
#         try:
#             ac.window(title=u"登录", auto_id="btnOK", control_type="Button").click()
#         except:
#             pass    

#     time.sleep(1)
#     if ac.window(title="登录错误", auto_id="BodyBounds", control_type="Pane").exists() == True:
#         err_win = ac.window(title="登录错误", auto_id="BodyBounds", control_type="Pane")
#         err_button = err_win.child_window(title="确认", auto_id="btnYes", control_type="Button")
#         err_button.wait('exists', timeout=5, retry_interval=1)
#         err_button.click()
#         if retry == True:
#             return try_login(False)

# try_login(True)
# uid = regedit.get_pc_id()
# print(uid)

# 关闭多余窗口
# app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
# ac = app.windows()
# if len(ac) > 1:
#     for i in ac:
#         print(i.automation_id())
#         if i.automation_id() != 'MDIMainForm':
#             i.close()
        # win = i.window(title_re="增值税发票税控开票软件")
        # print(i.window(title_re="增值税发票税控开票软件").exists())

# import wx
# import ContactUs
 
# class YPiao(ContactUs.MyFrame3):
#     #这里开始继承后对Virtual event handlers进行override,这个示例是对关于我们的菜单选择后进行了重载。
#     def make_sure( self, event ):
#         self.Close()
         
# # init the programe
# app = wx.App() #实例化APP，因为wxformbuilder只提供界面布局，所以需要我们自己对代码进行构架
# frame = YPiao(None) #frame的实例
# frame.Show();
 
# app.MainLoop() #wxpython的启动函数
# def months(date1,date2):
#     date1_splite = date1.split('-')
#     date2_splite = date2.split('-')
#     y = int(date2_splite[0]) - int(date1_splite[0])
#     m = int(date2_splite[1]) - int(date1_splite[1])
#     return y*12 + m

# ym = '2019-04-11'
# now = time.strftime("%Y-%m-01",  time.localtime())
# m = months(ym,now)
# print(m)
def post_link(link,post_data):
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
    try:
        req = requests.post(link, data=post_data, headers=headers)
        resp = req.json()
        return resp
    except:
        return {"code":500,"text":"网络异常，请求失败"}

ret = post_link('http://tinterface.hfxscw.com/interface.php?r=tax/tax-qk',{"credit_code":"91340100MA2NPN203B"})     

print(ret)

def tax_qk():
    app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",found_index=0)
    time.sleep(2)
    ac = app.window(title="税控发票开票软件")
    ac["登录"].click()
    if app['系统参数设置'].exists() == True:
        pass
    ac = app.window(class_name_re="WindowsForms10.Window.8.app",title="汇总信息")
    if ac.exists() == True:
        ac['确认'].click()

    time.sleep(2)
    app.SysMessageBox.wait('exists', timeout=20, retry_interval=1)
    # 关闭其它弹框
    if app.SysMessageBox.exists() == True:
        # 完成清卡操作
        mes = app.SysMessageBox.child_window(auto_id="lblMsg", control_type="Text")
        if '完成清卡' in mes.texts()[0]:
            print('清卡完成')
            pass

        app.SysMessageBox['确认'].click()

def getClipboardText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return(d).decode('GBK')

def exp_detail():
    app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",title_re="增值税发票税控开票软件")

    ac = app.window(class_name_re="WindowsForms10.Window.8.app")
    mon_win = ac.window(title="发票数据导出",auto_id="FPExport")
    # mon_win.print_control_identifiers()
    mon_sel = mon_win.window(title="月份", auto_id="cmbMonth")
    mon_sel.set_focus()

    # 检查导出的数据日期是否正确
    ok_btn = mon_win.child_window(title="确定",found_index=0, auto_id="btnOK", control_type="Button")
    if ok_btn.exists() == True:
        try:
            ok_btn.click()
        except:
            pass
        time.sleep(1)

    save_win = app.window(title='另存为')
    save_win.wait('exists', timeout=10, retry_interval=1)
    save_win.set_focus()
    pywinauto.keyboard.send_keys('^n^c')
    pwd = regedit.get_client_path()+"\\exp_file\\" + getClipboardText()
    save_win.ComboBox.Edit.set_text(pwd)
    pywinauto.keyboard.send_keys("^n%s")

# tax_qk()
# exp_detail()