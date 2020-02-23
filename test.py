from pywinauto.application import Application
import pywinauto
import time
import json
import configparser
import win32api
import win32con


# app = Application(backend='uia').connect(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1")
# ac = app.window(class_name="WindowsForms10.Window.8.app.0.1ed9395_r14_ad1")
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

app = Application(backend='uia').connect(class_name_re="WindowsForms10.Window.8.app",auto_id = "MDIMainForm")
ac = app.window(class_name_re="WindowsForms10.Window.8.app")
ac.set_focus()
export_window = ac.window(title="发票数据导出", auto_id="FPExport", control_type="Window")
export_window.window(title="确定", auto_id="btnOK", control_type="Button").click()


# app = Application(backend='uia').connect(title="SysMessageBox",class_name_re="WindowsForms10.Window.8.app")

# if app.window(title="SysMessageBox").exists() == True:
#     mes_win = app.window(title="SysMessageBox")
#     if mes_win.window(title="您选择的时间段没有数据", auto_id="lblMsg").exists() == True:
#         print('没有数据')
#         mes_win.window(title="确认", auto_id="btnYes", control_type="Button").click()
# time.sleep(1)
# print('1')
# app = Application(backend='uia').connect(title="另存为")
# print('2')
# save_win = app.window(title='另存为')
# print('3')
# save_win.set_focus()
# print('4')
# time.sleep(1)
# save_win.child_window(title="上一个位置", control_type="Button").click()
# print('5')