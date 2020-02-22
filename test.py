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

app = Application(backend='uia').connect(class_name="WindowsForms10.Window.8.app.0.15a303f_r12_ad1",title="税控发票开票软件")
ac = app.windows()[0]
ac.set_focus()