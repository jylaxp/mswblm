#-*-coding:gbk-*-
__author__ = 'Administrator'
import time
import win32com.client
import pythoncom
import sys
reload(sys)
sys.setdefaultencoding('gbk')

class DAO():
    def __init__(self, dbfile, systemfile):
        pythoncom.CoInitialize()
        self.lastLoginID = 0
        self.personnelInfoLastID = 0
        self.loginfoDNS = r'Provider=Microsoft.Jet.OLEDB.4.0;User ID=ikmdb;Data Source=%s\Loginfo.mdb;Persist Security Info=False;Jet OLEDB:System database=%s\System.mdw' % (dbfile, systemfile)
        self.netHouseDNS = r'Provider=Microsoft.Jet.OLEDB.4.0;User ID=ikmdb;Data Source=%s\NetHouseSer.mdb;Persist Security Info=False;Jet OLEDB:System database=%s\System.mdw' % (dbfile, systemfile)
        self.insiderDNS = r'Provider=Microsoft.Jet.OLEDB.4.0;User ID=ikmdb;Data Source=%s\Insider.mdb;Persist Security Info=False;Jet OLEDB:System database=%s\System.mdw' % (dbfile, systemfile)
        self.__getLastLoginID()
        self.__getPersonnelInfoLastID()

    def getWorkLoginInfo(self):
        #获取日志
        conn = win32com.client.Dispatch(r'ADODB.Connection')
        dns = self.loginfoDNS
        conn.Open(dns)
        sql = r'SELECT ID, SCardNumber, ComputerName, sCommand, sNote, sDate FROM [WorkLogInfo] WHERE ID>%s' % str(self.lastLoginID)
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql, conn, 1, 3)
        log = []
        while not rs.EOF:
            d = {}
            d['ID'] = int(rs('ID'))
            d['SCardNumber'] = str(rs('SCardNumber')).strip()
            d['ComputerName'] = str(rs('ComputerName')).strip()
            d['sCommand'] = str(rs('sCommand')).strip()
            d['sNote'] = str(rs('sNote')).strip()
            d['sDate'] = str(rs('sDate')).strip()
            log.append(d)
            self.lastLoginID += 1
            rs.MoveNext()
        rs.Close()
        conn.Close()
        return log

    def getInsiderinfo(self, insiderNumber):
        #获取会员信息
        conn = win32com.client.Dispatch(r'ADODB.Connection')
        dns = self.insiderDNS
        conn.Open(dns)
        sql = r"SELECT TOP 1 ID, InsiderNumber, TransactName, TransactTime, InsiderMoney FROM `insiderinfo` WHERE InsiderNumber='%s'" % insiderNumber
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql, conn, 1, 3)
        d = None
        if not rs.EOF:
            d = {}
            d['ID'] = int(rs('ID'))
            d['InsiderNumber'] = str(rs('InsiderNumber')).strip()
            d['TransactName'] = str(rs('TransactName')).strip()
            d['TransactTime'] = str(rs('TransactTime')).strip()
            d['InsiderMoney'] = float(str(rs('InsiderMoney')).strip())
        rs.Close()
        conn.Close()
        return d

    def getPersonnelInfos(self):
        #获取充值事件
        conn = win32com.client.Dispatch(r'ADODB.Connection')
        dns = self.netHouseDNS
        conn.Open(dns)
        sql = r'SELECT ID, ComputerName, Insider, InsiderNumber, SCardType, BeginTime, EndTime, YSMoney FROM `personnelinfo` WHERE ID>%s' % str(self.personnelInfoLastID)
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql, conn, 1, 3)
        L = []
        while not rs.EOF:
            d = {}
            d['ID'] = int(rs('ID'))
            d['ComputerName'] = str(rs('ComputerName')).strip()
            d['Insider'] = str(rs('Insider')).strip()
            d['InsiderNumber'] = str(rs('InsiderNumber')).strip()
            d['SCardType'] = str(rs('SCardType')).strip()
            d['BeginTime'] = str(rs('BeginTime')).strip()
            d['EndTime'] = str(rs('EndTime')).strip()
            d['YSMoney'] = float(str(rs('YSMoney')).strip())
            L.append(d)
            self.personnelInfoLastID += 1
            rs.MoveNext()
        rs.Close()
        conn.Close()
        return L

    def __getLastLoginID(self):
        conn = win32com.client.Dispatch(r'ADODB.Connection')
        dns = self.loginfoDNS
        conn.Open(dns)
        sql = r'SELECT Max(ID) AS ID FROM WorkLogInfo;'
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql, conn, 1, 3)
        if not rs.EOF:
            self.lastLoginID = int(rs('ID'))
        rs.Close()
        conn.Close()
        return self.lastLoginID

    def __getPersonnelInfoLastID(self):
        conn = win32com.client.Dispatch(r'ADODB.Connection')
        dns = self.netHouseDNS
        conn.Open(dns)
        sql = r'SELECT MAX(ID) AS ID FROM `personnelinfo`;'
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql, conn, 1, 3)
        if not rs.EOF:
            self.personnelInfoLastID = int(rs('ID'))
        rs.Close()
        conn.Close()
        return self.personnelInfoLastID
