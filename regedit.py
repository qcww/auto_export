import win32con
import win32api
import winreg
import os

def add_auto_run(path):
    # zdynames = os.path.basename(__file__)     # 当前文件名的名称如：newsxiao.py
    name = 'AutoExport'      # 获得文件名的前部分,如：newsxiao
    # path = os.path.abspath(os.path.dirname(__file__))+'\\'+zdynames # 要添加的exe完整路径如：
    # 注册表项名
    KeyName = 'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    # 异常处理
    try:
        key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER,  KeyName, 0,  win32con.KEY_ALL_ACCESS)
        win32api.RegSetValueEx(key, name, 0, win32con.REG_SZ, path)
        win32api.RegCloseKey(key)
    except:
        return False
    return True

# add_auto_run("D:\\phpstudy\\PHPTutorial\\WWW\\auto_export\\main.exe /background")

def set_client_path():
    reg_root = win32con.HKEY_LOCAL_MACHINE
    reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AutoExport.exe"
    reg_flags = win32con.WRITE_OWNER|win32con.KEY_WOW64_64KEY|win32con.KEY_ALL_ACCESS

    try:
        #直接创建（若存在，则为获取）
        key, _ = win32api.RegCreateKeyEx(reg_root, reg_path, reg_flags)

        #设置项
        config_path = os.path.abspath(os.path.dirname(__file__))
        win32api.RegSetValueEx(key, "Path", 0, win32con.REG_SZ, config_path)

        #关闭
        win32api.RegCloseKey(key)
        return True
    except:
        return False


def get_app_path(app_path,app_name):
    app_path = app_path
    if app_path == "":
        try:
            reg_key = winreg .OpenKey(winreg .HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\fwkp.exe")
            app_path_all, _ = winreg.QueryValueEx(reg_key, "Path")

            path_start = app_path_all.find('开票软件')
            app_path = app_path_all[0:int(path_start)+5] + app_name
        except:
            app_path = ''
    else:
        app_path = app_path + app_name
    if app_path != '' and os.path.exists(app_path):
        return app_path
    return ''


def get_client_path():
    try:
        reg_key = winreg .OpenKey(winreg .HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AutoExport.exe")
        app_path, _ = winreg.QueryValueEx(reg_key, "Path")
    except:
        app_path = ''

    return app_path