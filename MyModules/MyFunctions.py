import os, sys
from PyQt5 import QtCore, QtWidgets, QtGui, QtSql
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtSql import (QSqlDatabase,  QSqlQueryModel, QSqlQuery)

from fnmatch import fnmatch
import socket
import csv 
import subprocess
import inspect
from time import time, sleep, monotonic
import uuid


# 獲取Mac地址
def get_mac_address():
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:]
    return ":".join([mac[e:e+2] for e in range(0,11,2)])

#取得主機名稱
def get_hostname():
    try:
        hostname = os.environ['COMPUTERNAME']
        return hostname
    except KeyError:
        try:
            hostname = os.uname().nodename
            return hostname
        except AttributeError:
            return "Unknown Hostname"
#取得主機ip
def get_ip():
    try:
        ip = subprocess.check_output(['hostname', '-I']).decode('utf-8').strip()
        return ip
    except subprocess.CalledProcessError:
        return "Unknown IP"
        
#判斷字串是否可以轉成浮點數
def isnumber(aString):
    try:
        if aString!='nan':
            float(aString)
            return True
        else:
            return False
    except:
        return False
#轉成字串, 並將None轉為''空字串
def xstr(s):
    if s is None:
        return ''
    else:
        return str(s)

def wait_sec(delaySec=0.0,start_monotonic=''):
    try:
        wt=0
        if start_monotonic=='':
            st=monotonic()
        else:
            st=start_monotonic            
        while(wt<delaySec):
            wt=monotonic()-st
            QtWidgets.QApplication.processEvents()
            # if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈
            sleep(0.001)
    except Exception as e:
        print(str(e))