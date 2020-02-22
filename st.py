import winreg 
reg_key=winreg .OpenKey(winreg .HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\fwkp.exe")
app_path_all, type = winreg.QueryValueEx(reg_key, "Path")

path_start = app_path_all.find('开票软件')
print(path_start,app_path_all[0:19])