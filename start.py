#-*- coding: utf-8 -*-

from upgrade import update
import subprocess
import json
import configparser
import os

class start:
    def __init__(self):
        self.config_file = './config/config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8')

    def upgrade(self):
        upgrade = update()
        if upgrade.check_update() == True:
            upgrade.download()

    def run_app(self):
        path = self.config['client']['path']
        name = self.config['client']['name']
        os.chdir(path)
        subprocess.call(name)

app = start()
app.upgrade()
app.run_app()