#-*- coding: utf-8 -*-

from upgrade import update
import subprocess
import json
import configparser
import os
import regedit
import sys
import re

class start:
    def __init__(self):
        self.config_file = "./config/config.ini"
        if os.path.exists(self.config_file) == False:
            self.config_file = regedit.get_client_path()+"\\config\\config.ini"
        else:
            regedit.set_client_path()
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')

        if '.exe' in sys.argv[0]:
            # 设置开机启动
            pp = re.findall(r'.[^\\]*\.exe',sys.argv[0])
            ename = pp[0].replace('\\','')
            main_app = regedit.get_client_path() + '\\' + ename + " /autorun"
            if regedit.add_auto_run(main_app) == False:
                print('设置开机启动失败')
            else:
                print('设置开机启动成功')    

    def upgrade(self):
        upgrade = update()
        if upgrade.check_update() == True:
            upgrade.download()

    def run_app(self):
        path = regedit.get_client_path() +"\\main\\"+ self.config['client']['name']

        subprocess.call(path)

app = start()
app.upgrade()
app.run_app()