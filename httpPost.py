#-*-coding:utf-8-*-
__author__ = 'Administrator'
import urllib
import urllib2

class HttpPost():
    def __init__(self):
        pass

    def post(self, url, data):
        values = {}
        values['data'] = data
        data = urllib.urlencode(values)
        request = urllib2.Request(url, data)
        response = urllib2.urlopen(request)
        res = response.read().strip()
        return res