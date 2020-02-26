from pywinauto.application import Application
import pywinauto
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

# app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app")
# ac = app.window(class_name_re="WindowsForms10.Window.8.app")
# mon_win = ac.window(title="发票数据导出")
# # mon_win.print_control_identifiers()
# mon_sel = mon_win.window(title="月份", auto_id="cmbMonth", control_type="ComboBox")
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
def upload_file(file_name,upload_link,data):
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
    file_name = './exp_file/'+file_name
    if os.path.exists(file_name):
        print('fff')
    files = {'upfile': open(file_name, 'rb')}
    res = requests.post(upload_link, data=data, files=files, headers=headers)
    return res

re = upload_file('增值税专普发票数据导出20200225.xlsx','http://tinterface.hfxscw.com/interface.php?r=tax/invoice-upload',{"corpid":"71","period":"202001","submit":"1","tax_import":"1"})
print(re.json())