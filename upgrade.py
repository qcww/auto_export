from urllib import request
import configparser
import requests
import wx
import htool
import json
import os
 
class update:
    def __init__(self):
        self.config_file = './config/config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file,encoding='utf-8') 

    def check_update(self):
        try:
            data = requests.get(self.config['link']['client_update_url'])
            parse = data.json()
            parse_data = json.loads(parse['data'])
            self.dist_version = parse_data['version']
            
            if self.config['client']['version'] != self.dist_version:
                ok = htool.ask('升级提示','系统检测到有新的版本,是否立即更新')
                if ok == 6:
                    return True
        except:
            htool.warn('提示','网络连接错误,请稍后重试')
            return False
        # 通过接口请求信息 判断是否最新客户端
        return False

    def download(self):
        base_url = self.config['link']['dw_url']
        #使用下载函数下载视频并调用进度函数输出下载进度
        client = self.config['client']

        main_file = client['path']+client['name']
        if os.path.exists(main_file):
            os.remove(main_file)
        try:
            main_file = client['path']+client['name']
            if os.path.exists(main_file):
                os.remove(main_file)
            request.urlretrieve(url=base_url,filename=main_file,reporthook=self.report,data=None)
        except:
            pass
        

    #下载进度函数
    def report(self,a,b,c):
        '''
        a:已经下载的数据块
        b:数据块的大小
        c:远程文件的大小
        '''
        per = 100.0 * a * b / c
        if per > 100:
            per = 100
            self.config.set('client','version',self.dist_version)
            with open(self.config_file,'w') as f:
                self.config.write(f)