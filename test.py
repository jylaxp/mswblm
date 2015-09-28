#-*-coding:utf-8-*-
__author__ = 'Administrator'

import dataAccess
import httpPost

import win32serviceutil
import win32service
import win32event
import time
import json
import traceback
#import sys
#reload(sys)
#sys.setdefaultencoding('gbk')

class mswblmppyService():
    _svc_name_ = "mswblmppy"
    _svc_display_name_ = "mswblmppy"
    _svc_description_ = "mswblmppy Service"

    def __init__(self, args):
        self.logger = self._getLogger()
        self.isAlive = True
        self.payDic = {}

        self.dbfile = None
        self.systemfile = None
        self.netbar = None
        self._config()

        self.dao = dataAccess.DAO(self.dbfile, self.systemfile)
        self.post = httpPost.HttpPost()
        #self.url = r'http://www.mswblm.com/wang/index.php?g=Wang&m=User&a=bar_test'
        self.url = r'http://www.mswblm.com/wang/index.php?g=Wang&m=Test&a=bar_test'

    def _getLogger(self):
        import logging
        import os
        import inspect
        logger = logging.getLogger('[mswblmppy]')
        this_file = inspect.getfile(inspect.currentframe())
        dirpath = os.path.abspath(os.path.dirname(this_file))
        handler = logging.FileHandler(os.path.join(dirpath, "mswblmppy.log"))
        formatter = logging.Formatter('%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
        return logger

    def _config(self):
        import ConfigParser
        config = ConfigParser.ConfigParser()
        config.read(r"C:\WINDOWS\mswblm.ini")
        self.dbfile = config.get("mswblm", "AccessPath")
        self.systemfile = config.get("mswblm", "SystemPath")
        self.netbar = config.get("mswblm", "NetBar").decode('gbk').encode('utf-8')
        print self.netbar

    def SvcDoRun(self):
        import time
        self.logger.debug("svc do run....")
        while self.isAlive:
            try:
                self.logger.debug("before process")
                self._process()
                self.logger.debug("after process")
            finally:
                time.sleep(1)

    def SvcStop(self):
        self.logger.debug("svc do stop....")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.isAlive = False

    def _process(self):
        #self._eventHanlde()
        self._payEventHandle()


    def _eventHanlde(self):
        logs = self.dao.getWorkLoginInfo()
        LL = None
        for log in logs:
            if LL is None:
                LL = []
            if log['sCommand'].find(u'刷卡计费') > -1 or log['sCommand'].find(u'客户端登录') > -1 or log['sCommand'].find(u'结帐') > -1:
                info = self.dao.getInsiderinfo(log['SCardNumber'])
                if info:
                    d = {}
                    d['netbar'] = self.netbar
                    #d['idcard'] = info['InsiderNumber']
                    d['idcard'] = '510107199007114395'
                    d['pcnum'] = log['ComputerName']
                    d['money'] = info['InsiderMoney']
                    if log['sCommand'].find(u'刷卡计费') > -1:
                        d['type'] = 1
                    elif log['sCommand'].find(u'客户端登录') > -1:
                        d['type'] = 2
                    elif log['sCommand'].find(u'结帐') > -1:
                        d['type'] = 3
                    else:
                        d['type'] = 3
                    LL.append(d)

        if LL:
            data = json.dumps(LL, ensure_ascii=False)
            self.logger.debug(data)
            self.post.post(self.url, data)

    def _payEventHandle(self):
        payEvents = self.dao.getPersonnelInfos()
        for evt in payEvents:
            if evt['ComputerName'].find(u'会员充值') > -1:
                e = self.payDic.get(evt['InsiderNumber'])
                if e is None:
                    self.payDic[evt['InsiderNumber']] = self._getPayObject(evt)
                else:
                    e['cash'] += evt['YSMoney']
                    e['time'] = int(time.time())

        dataList = None
        for k, v in self.payDic.iteritems():
            if dataList is None:
                dataList = []
            now = int(time.time())
            if now - v['time'] > 10:
                dataList.append(v)

        if dataList is None:
            return

        dl = None
        for d in dataList:
            self.payDic.pop(d['idcard'])
            info = self.dao.getInsiderinfo(d['idcard'])
            if info:
                if dl is None:
                    dl = []
                d['money'] = info['InsiderMoney']
                d['idcard'] = '510107199007114395'
                dl.append(d)

        if dl:
            data = json.dumps(dl, ensure_ascii=False)
            self.logger.debug(data)
            self.post.post(self.url, data)

    def _getPayObject(self, evt):
        dd = {}
        dd['netbar'] = self.netbar
        dd['idcard'] = evt['InsiderNumber']
        dd['pcnum'] = evt['ComputerName']
        dd['cash'] = evt['YSMoney']
        dd['type'] = 4
        dd['money'] = 0
        dd['time'] = int(time.time())
        return dd


msw = mswblmppyService(1)
msw.SvcDoRun()
