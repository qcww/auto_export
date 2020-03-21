import win32com.client

def check_usb(key):
    wmi = win32com.client.GetObject ("winmgmts:")
    usb_insert = False
    for usb in wmi.InstancesOf ("Win32_USBHub"):
        # print(usb.Dependent)
        if key in usb.DeviceID:
            return True
    return False