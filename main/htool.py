#-*- coding: utf-8 -*-
import wx
import win32api,win32con

def alert(title,content):
    win32api.MessageBox(0, content, title,win32con.MB_ICONASTERISK)

def warn(title,content):
    win32api.MessageBox(0, content, title,win32con.MB_ICONWARNING)

def ask(title,content):
    return win32api.MessageBox(0, content, title,win32con.MB_YESNO)