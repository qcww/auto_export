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
#导入SQLite驱动：
import sqlite3

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


#连接到SQlite数据库
#数据库文件是test.db，不存在，则自动创建
conn = sqlite3.connect('ff.db')
#创建一个cursor：
cursor = conn.cursor()
#执行一条SQL语句：创建user表
cursor.execute('create table user(id varchar(20) primary key,name varchar(20))')
#插入一条记录：
cursor.execute('insert into user (id, name) values (\'1\', \'Michael\')')
#通过rowcount获得插入的行数：
print(cursor.rowcount) #reusult 1
#关闭Cursor:
cursor.close()
#提交事务：
conn.commit()
#关闭connection：
conn.close()