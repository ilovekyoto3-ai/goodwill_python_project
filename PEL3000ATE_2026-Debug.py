#-----------------------------------------------------------------------------------------------------------------------
__version__='0.0.1.23'               #軟體版本
__author__='Vincent'
__description__=" ATE測試系統"
__product_name__='ATE_PEL-3000'
WTitle="PEL-3XXX ATE測試系統" #視窗標題
appVer='0.1.0'                      #顯示軟體版本
url_info = "http://172.16.13.251:8180/ATE_PSU/program_info.json"

#----------------------------------------------------------------------------------------------------------------------

# from asyncio.windows_events import NULL
# from logging import exception
# from pickle import TRUE
# from posixpath import split
# from re import S
# from telnetlib import DM
# from tkinter import HORIZONTAL
# from turtle import delay
# from numpy import typename

from time import time, sleep, monotonic
from MyModules import MyInstr, MyDB
from Ui_Main_Test import Ui_frmMainTest
from Ui_frmTRecord import Ui_frmTRecord
from Ui_frmMaintainDB import Ui_frmMaintainDB
from PyQt5 import QtCore, QtWidgets, QtGui, QtSql
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtSql import (QSqlDatabase,  QSqlQueryModel, QSqlQuery)

import sqlite3                   
import pathlib
from fnmatch import fnmatch

# import socket
import uuid
import csv, sys
import subprocess
import inspect
import pymysql
import os
from os import path, access

TopMargin=30
leftMargin=10
AteDB=MyDB.SQLiteDB()


#設定待測機或測試設備的物件-----------------------------------------------------------------------
DUT_RS232=MyInstr.PEL()
DUT_USB=MyInstr.PEL()
DUT_GPIB=MyInstr.PEL()
DMM=MyInstr.DMM()
HVM=MyInstr.vitrek_4700()
PCS=MyInstr.PCS1000()
DIO=MyInstr.MEGA2560_GPT()
PSW=MyInstr.DC_PS()


#------------------------------------------------------------------------------------------------

TESTNUM=3

#自訂顏色常數
color_Magenta=QColor(255,0,255)
color_Darkviolet=QColor(148,0,21)
color_Violet=QColor(238,130,238)

   
# 獲取Mac地址
def get_mac_address():
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:]
    return ":".join([mac[e:e+2] for e in range(0,11,2)])
#判斷字串是否可以轉成浮點數
def isnumber(aString):
    try:
        float(aString)
        return True
    except:
        return False
#轉成字串, 並將None轉為''空字串
def xstr(s):
    if s is None:
        return ''
    else:
        return str(s)
#關閉主視窗        
def MainClose():
    window.close()

def showTestResult():
    try:
        window.TRecord.TestDocNo=window.MainTest.cboReportNo.currentText()
        window.TRecord.TestModel=window.MainTest.ltxtDUTModel.text()
        window.TRecord.DBPth=window.DBPth
        window.TRecord.show()
        window.TRecord.showTestData()
        window.TRecord.adjColWidth()
        if window.MainTest.x()<=10:
            window.MainTest.move(400,150)
        window.MainTest.show()
        window.MainTest.activateWindow()
    except Exception as e:
        QMessageBox.critical(window,'showTestResult錯誤',f'Error:{str(e)}')


class MainWindow(QtWidgets.QMainWindow):
    count=0
    def __init__(self, parent=None):
        super(MainWindow,self).__init__(parent)
              
        self.mdi=QMdiArea()
        self.setCentralWidget(self.mdi)
        bar=self.menuBar()
        file=bar.addMenu('檔案')
        view=bar.addMenu('檢視')
        tools=bar.addMenu('工具')
        #新增子選單
        file.addAction('結束')
        view.addAction('測試主視窗')
        view.addAction('測試記錄視窗')
        view.addAction('資料維護視窗')
        tools.addAction('上傳測試記錄')
        tools.addAction('下載雲端腳本')
        file.triggered[QtWidgets.QAction].connect(self.windowaction)
        view.triggered[QtWidgets.QAction].connect(self.windowaction)
        tools.triggered[QtWidgets.QAction].connect(self.windowaction)
        #設定主視窗的標題
        self.setWindowTitle(f'{WTitle} (版本:{appVer})')  
        self.showMaximized()
        
        # try:
        #     self.MainTest=MainTest()
        #     self.mdi.addSubWindow(self.MainTest)
        #     self.DBPth=self.MainTest.DBPth         
        #     self.TRecord=frmTestRecord()
        #     self.TRecord.DBPth=self.DBPth
        #     self.mdi.addSubWindow(self.TRecord)
        #     self.MaintainDB=frmMaintainDB()
        #     self.MaintainDB.DBPth=self.DBPth
        #     self.mdi.addSubWindow(self.MaintainDB)
        # except Exception as e:
        #     QMessageBox.critical(self,'MainWindow __init__',f'Error:{str(e)}')
    
    def onQApplicationStarted(self):
        try:
            print ('started')
            self.MainTest=MainTest()
            self.mdi.addSubWindow(self.MainTest)
            self.DBPth=self.MainTest.DBPth         
            self.TRecord=frmTestRecord()
            self.TRecord.DBPth=self.DBPth
            self.mdi.addSubWindow(self.TRecord)
            self.MaintainDB=frmMaintainDB()
            self.MaintainDB.DBPth=self.DBPth
            self.mdi.addSubWindow(self.MaintainDB)


            self.MainTest=MainTest()
            self.mdi.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            self.mdi.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            self.mdi.addSubWindow(self.MainTest) 
            self.MainTest.show()
            QtWidgets.QApplication.processEvents()
            self.MainTest.appStarted()

        except Exception as e:
            QMessageBox.critical(self,f"{inspect.currentframe().f_code.co_name}",f'Error:{str(e)}')


    def setDBPth(self,s):
        self.DBPth=s
    
    def windowaction(self,q):
        print('windowaction Triggered')
        if q.text()=='結束':
            self.close()
        if q.text()=='測試主視窗':
            self.MainTest.show()

        if q.text()=='測試記錄視窗':
            self.TRecord.show()
            self.TRecord.DBPth=self.DBPth
            self.TRecord.TestDocNo=self.MainTest.cboReportNo.currentText()
            self.TRecord.TestModel=self.MainTest.ltxtDUTModel.text()
            self.TRecord.showTestData()
            self.TRecord.adjColWidth()

        if q.text()=='資料維護視窗': 
            self.MaintainDB.show()
            self.MaintainDB.DBPth=self.DBPth           
            # if self.MainTest.ltxtInspector.text().lower().find('admin')>=0:            
            #     self.MaintainDB.show()
            #     self.MaintainDB.DBPth=self.DBPth
            # else:
                # QMessageBox.critical(self,'維護者不符','請在測試主視窗的檢測者輸入維護者名稱.')
        
        if q.text()=='上傳測試記錄':
            self.MainTest.show()
            self.MainTest.btnStop.setEnabled(True)
            self.MainTest.StopTest=False
            self.MainTest.UploadTestData()
            self.MainTest.btnStop.setEnabled(False)
        
        if q.text()=='下載雲端腳本':
            self.MainTest.show()
            self.MainTest.btnStop.setEnabled(True)
            self.MainTest.StopTest=False
            self.MainTest.DownloadTestScript()
            self.MainTest.getTestScript()
            self.MainTest.selectScript()
            self.MainTest.btnStop.setEnabled(False)
            


    
    def closeEvent(self, event):
        try:
            subws=self.findChildren(QMdiSubWindow)            
            for sw in subws:
                sw.close()
                print(f"{sw.objectName()} closed.")
        except Exception as e:
            QMessageBox.critical(self,'Main window close error',f'Error:{str(e)}')


#為格式化顥示, 自訂義的SqlTableModel
class CustomSqlModel(QtSql.QSqlTableModel):
    def __init__(self, parent=None):
        QtSql.QSqlTableModel.__init__(self, parent=parent)
        
        

    def data(self, index, role):
        if role == QtCore.Qt.BackgroundRole: #處理背景           
            if str(self.record(index.row()).value('Connected')).lower()=='false' or str(self.record(index.row()).value('Connected'))=='0':
                return QtGui.QBrush(QtGui.QColor(255, 255, 200))
            elif str(self.record(index.row()).value('Connected')).lower()=='true' or str(self.record(index.row()).value('Connected'))=='1':
                return QtGui.QBrush(QtGui.QColor(180, 255, 180))
            else:
                return QtGui.QBrush(QtGui.QColor(200, 200, 200))
        if role == QtCore.Qt.ForegroundRole:   #處理前景          
            if str(self.record(index.row()).value('Connected')).lower()=='false' or str(self.record(index.row()).value('Connected'))=='0':
                return QtGui.QBrush(QtCore.Qt.red)
            elif str(self.record(index.row()).value('Connected')).lower()=='true' or str(self.record(index.row()).value('Connected'))=='1':
                return QtGui.QBrush(QtCore.Qt.black)
            else:
                return QtGui.QBrush(QtCore.Qt.black)            
        return QtSql.QSqlTableModel.data(self, index, role)

#為格式化顥示, 自訂義的SqlQueryModel
class CustomSqlQuery(QtSql.QSqlQueryModel):
    def __init__(self, parent=None):
        QtSql.QSqlQueryModel.__init__(self, parent=parent)
        self.TestRow=-1          #正在測試的列
        
    def data(self, index, role):  
        if role == QtCore.Qt.BackgroundRole: #處理背景           
            if index.row()==self.TestRow and index.column() not in [3,4,5]:
                return QtGui.QBrush(QtGui.QColor(180, 255, 180))
            else:
                # print(f"self.TestRow:{self.TestRow}, index.row:{index.row()}")
                return QtGui.QBrush(QtGui.QColor(240, 240, 240))

        # if (role == QtCore.Qt.ForegroundRole and index.column() >=13):
        if (role == QtCore.Qt.ForegroundRole):    #處理前景        
                if str(self.record(index.row()).value('判定')).lower()=='fail' or str(self.record(index.row()).value('判定'))=='ng':
                    return QtGui.QBrush(QtCore.Qt.red)
                elif str(self.record(index.row()).value('判定')).lower()=='pass' or str(self.record(index.row()).value('判定'))=='ok':
                    return QtGui.QBrush(QtCore.Qt.blue)
                else:
                    return QtGui.QBrush(QtCore.Qt.black)                
        return QtSql.QSqlQueryModel.data(self, index, role)

#建立自訂的GroupBox類別, 可以回應mouse雙擊, 全選或全不選內含的選項
class MyQGroupBox(QGroupBox):
    def __init__(self, parent=None):
        QGroupBox.__init__(self, parent=parent)
        self.allchecked=False

    def mouseDoubleClickEvent(self, event):
        chkBoxs=self.findChildren(QCheckBox)
        self.allchecked=False
        for chkB in chkBoxs:
            if chkB.isChecked():
                self.allchecked=True
        self.allchecked=not self.allchecked
        for chkB in chkBoxs:
            if chkB.isEnabled():
                chkB.setChecked(self.allchecked)

#測試記錄欄位名稱    
class recordHeader():
    col_No_1='No_1'
    col_No_2='No_2'
    col_No_3='No_3'
    col_ItemName='測項名稱'
    col_ItemName_2t='次項名稱'
    col_ItemName_3t='次次項名稱'
    col_CondName_1 ='條件1名稱'
    col_Cond_1 ='條件1值'
    col_CondUnit_1 ='條件1單位'
    col_CondName_2 ='條件2名稱'
    col_Cond_2 ='條件2值'
    col_CondUnit_2 ='條件2單位'
    col_CondName_3 ='條件3名稱'
    col_Cond_3 ='條件3值'
    col_CondUnit_3 ='條件3單位'
    col_Expect ='期望值'
    col_ExpectUnit ='單位'
    col_LowLimit ='下限值'
    col_UpLimit ='上限值'
    col_Measure_Main ='主量測值'
    col_Judge ='判定'
    col_Measure_1 ='次量測值1'
    col_Measure_2 ='次量測值2'
    col_Deviation ='主偏差值'
    col_Deviation_Expect ='偏差期望值'
    col_Deviation_LowLimit ='偏差下限'
    col_Deviation_UpLimit ='偏差上限'
    col_Cond_Fixture='治具條件'
    col_Cond_DUT='條件_DUT'
    col_Comment ='備註'  
    col_TDatetime='測試日期時間'              

#顯示測試記錄視窗
class frmTestRecord(QMdiSubWindow, Ui_frmTRecord):
    StopTest=False
    TestDocNo=''
    TestModel=''
    DBPth=''
    def __init__(self, parent=None):
        super(frmTestRecord,self).__init__(parent)        
        self.setupUi(self)
        self.db =  QSqlDatabase.addDatabase('QSQLITE')              #建立資料庫物件
        self.model = CustomSqlQuery()
        self.settings = QSettings()
        self.FPth=self.settings.value('defaultFilePath')

    def resizeEvent(self, event): 
        super().resizeEvent(event)       
        width, height = event.size().width(), event.size().height()
        if not self.isMinimized():
            self.tbview.setGeometry(QtCore.QRect(10, 30, width-20, height-40))
        

    def showTestData(self): 
        subProcName='showTestData'
        try:          
            self.tbview.setWindowTitle("QSqlQueryModel")
            self.db =  QSqlDatabase.addDatabase('QSQLITE')          #建立資料庫物件
            # 设置数据库名称
            #self.DBPth=window.MainTest.ledDBPath.text()
            self.db.setDatabaseName(self.DBPth)
            if not self.db.open():  self.db.open()
            # 判断是否打开
            if not self.db.open():            
                return False
            
            #self.model = CustomSqlQuery()
            self.setWindowTitle(f'測試編號:{self.TestDocNo}')
            

            query_str = f"""select ItemNo as No_1, SubItemNo as No_2, TestPointNo as No_3 
                , ItemName as 測項名稱, ItemName_2t as 次項名稱, ItemName_3t as 次次項名稱 
                , CondName_1 as 條件1名稱, Cond_1 as 條件1值, CondUnit_1 as 條件1單位
                , CondName_2 as 條件2名稱, Cond_2 as 條件2值, CondUnit_2 as 條件2單位
                , CondName_3 as 條件3名稱, Cond_3 as 條件3值, CondUnit_3 as 條件3單位
                , Expect as 期望值, ExpectUnit as 單位, LowLimit as 下限值, UpLimit as 上限值
                , Measure_Main as 主量測值
                , Judge as 判定
                , Measure_1 as 次量測值1, Measure_2 as 次量測值2
                , Deviation as 主偏差值, Deviation_Expect as 偏差期望值
                , Deviation_LowLimit as 偏差下限, Deviation_UpLimit as 偏差上限
                , Comment as 備註 , TDatetime as 測試日期時間, cond_Fixture as 治具條件, cond_DUT as 條件_DUT
                from TestRecord where TestDocNo='{self.TestDocNo}' """
            self.model.setQuery(query_str, db=self.db)
            self.tbview.setModel(self.model)
        except Exception as e:
            QMessageBox.critical(self,'錯誤',f"{subProcName} Error: \n{str(e)}")



    def adjColWidth(self):
        self.tbview.setColumnWidth(0,20)    #No_1 
        self.tbview.setColumnWidth(1,20)    #No_2 
        self.tbview.setColumnWidth(2,20)    #No_3 

        self.tbview.setColumnWidth(3,60)    #測項名稱   
        self.tbview.setColumnWidth(4,60)    #次項名稱
        self.tbview.setColumnWidth(5,80)    #次次項名稱
        self.tbview.setColumnWidth(6,60)    #條件1名稱
        self.tbview.setColumnWidth(7,60)    #條件1值
        self.tbview.setColumnWidth(8,40)    #條件1單位
        self.tbview.setColumnWidth(9,60)    #條件2名稱
        #self.tbview.hideColumn(9)           
        self.tbview.setColumnWidth(10,50)   #條件2值
        self.tbview.setColumnWidth(11,50)   #條件2單位
        self.tbview.setColumnWidth(12,40)   #條件3名稱
        #self.tbview.hideColumn(12)
        self.tbview.setColumnWidth(13,50)   #條件3值
        #self.tbview.hideColumn(13)
        self.tbview.setColumnWidth(14,50)   #條件2單位
        #self.tbview.hideColumn(14)       
        self.tbview.setColumnWidth(15,60)   #期望值
        self.tbview.setColumnWidth(16,50)   #期望值單位
        self.tbview.setColumnWidth(17,60)   #下限
        #self.tbview.hideColumn(14)
        self.tbview.setColumnWidth(18,60)   #上限
        self.tbview.setColumnWidth(19,80)   #主量測值
        self.tbview.setColumnWidth(20,40)   #判定
        self.tbview.setColumnWidth(21,60)
        self.tbview.setColumnWidth(22,60)
        self.tbview.setColumnWidth(23,40)
        self.tbview.setColumnWidth(24,40)
        self.tbview.setColumnWidth(25,40)
        self.tbview.setColumnWidth(26,60)
        self.tbview.setColumnWidth(27,80)   #備註
        self.tbview.setColumnWidth(28,140)   #測試日期時間
        self.tbview.setColumnWidth(29,80)  #治具條件 
            
        self.CellSpan()
    
    def CellSpan(self):
        # self.tbview.setSpan(0,2,3,1)        
        try:
            self.tbview.verticalScrollBar().setDisabled(True)
            iname=['測項名稱','次項名稱','次次項名稱']
            for i in range(self.model.rowCount()):  #從第一列掃到最後一列
                 for k in range(0,len(iname)):
                    self.tbview.setSpan(i,k+3,1,1)

            for k in range(0,len(iname)):
                print('')
                j=1
                j1=0
                CV=''       #目前值
                LV=''       #最後值     
                for i in range(self.model.rowCount()):  #從第一列掃到最後一列
                    r=self.model.record(i)
                    CV=''
                    for l in range(0,k+1):
                        CV+=str(r.value(iname[l])).strip()                
                    
                    if CV==LV:
                        # print(j, r.value(iname[k]), LV)
                        j+=1                
                    else:
                        if j>1:
                            try:
                                self.tbview.setSpan(j1,k+3,j,1)
                                # print(f'合併, 欄位:{LV}, 閏始列:{j1}, 列數:{j}')
                                # QApplication.processEvents()
                            except Exception as e:
                                print(f'合併, 欄位:{LV}, 閏始列:{j1}, 列數:{j}發生錯誤:\n{str(e)}')
                            finally:
                                j=1
                        j1=i
                    LV=CV
        except Exception as e:
            print(f"{inspect.currentframe().f_code.co_name} Error: {str(e)}")
        finally:
            self.tbview.verticalScrollBar().setEnabled(True)

    #匯出表格資料到CSV檔
    def table_export(self):
        try:
            if self.FPth=='': self.FPth=path.dirname(self.DBPth)
            self.FPth=f'{self.FPth}\\{self.TestModel} {self.TestDocNo}.csv'
            File_Name, File_Type=QFileDialog.getSaveFileName(self, "匯出測試資料", self.FPth, "csv Files (*.csv)")
            self.FPth=path.dirname(File_Name)
            print(File_Name, File_Type)
            if File_Name!='':
                rowdata = []
                with open(File_Name, 'w',newline='',encoding='big5') as stream:
                    writer = csv.writer(stream)
                    col_count=self.model.columnCount()
                    for c in range(col_count):
                        item = str(self.model.headerData(c,Qt.Horizontal) ) 
                        if item is not None:
                            rowdata.append(item)
                        else:
                            rowdata.append('')
                    writer.writerow(rowdata)
                    for row in range(self.model.rowCount()):
                        rowdata.clear()
                        for column in range(self.model.columnCount()):
                            item = self.model.record(row).value(column)
                            if item is not None:
                                rowdata.append(item)
                            else:
                                rowdata.append('')
                        writer.writerow(rowdata)
            self.settings.setValue('defaultFilePath',self.FPth)
        except Exception as e:
            QMessageBox.critical(self,'table_export 錯誤', f"{str(e)}")

#資料庫資料維護視窗
class frmMaintainDB(QMdiSubWindow, Ui_frmMaintainDB):
    DBPth=''    
    def __init__(self, parent=None):
        super(frmMaintainDB,self).__init__(parent)        
        self.setupUi(self)

        

        self.settings=QSettings("Maintain")
        self.FPth=self.settings.value('defaultFilePath')
        self.table1.setColumnCount(1)
        self.table1.setSelectionBehavior(QTableWidget.SelectItems)#設為單格選取
        self.table1.setSelectionMode(QAbstractItemView.ExtendedSelection)#設可選取多個儲存格
        self.table1.setHorizontalHeaderLabels(["欄位1"]) #設欄位表頭
        self.btnDelete.clicked.connect(self.table_delete)
        self.btnInsert.clicked.connect(self.table_insert)
        self.btnUpdate.clicked.connect(self.table_update)
        # self.btnQuery.clicked.connect(self.table_query)
        self.btnImport.clicked.connect(self.table_import)
        self.btnExport.clicked.connect(self.table_export)
        self.cboTable.currentIndexChanged.connect(self.on_cboTable_currentIndexChanged)
        self.cboUpScript.currentIndexChanged.connect(self.oncboUpScriptChanged)
        
        self.cboTable.addItem('治具路徑校正值')
        self.cboTable.addItem('機種測試腳本')
        self.cboTable.addItem('測試腳本')
        self.cboTable.addItem('治具路徑控制表')
        self.cboTable.addItem('治具埠腳位')
        self.cboTable.addItem('測試腳本重號檢查')
        
        self.FillScriptSelector()


    def resizeEvent(self, event): 
        super().resizeEvent(event)       
        width, height = event.size().width(), event.size().height()
    
        if not self.isMinimized():        
            # self.verticalLayout_4.setGeometry(QtCore.QRect(10, 30, width-10, height-40))
            self.groupBox_1.move(width-230,40)
            self.groupBox_2.move(width-230,170)
            self.groupBox_3.move(width-230,515)            
            self.table1.setGeometry(QtCore.QRect(10, 40, width-250, height-60))
             
    #     for a in [QLabel, QLineEdit, QComboBox,QPushButton]:
    #         Objs=self.findChildren(a)
    #         for O in Objs:
    #             x=O.x()
    #             y=O.y()
    #             w=O.width()
    #             O.move(width-w-10,y)

    def on_cboTable_currentIndexChanged(self):        
        self.on_btnQuery_clicked()
        
    #舊版程序名稱table_query, 0.1.40後改為on_btnQuery_clicked槽信號名稱
    @pyqtSlot()
    def on_btnQuery_clicked(self):
        rsFix=[]
        cols=0
        try:
            self.DBPth=window.DBPth
            self.table1.setColumnCount(1)
            self.table1.setRowCount(0)
            AteDB.ConnectSQLite(self.DBPth)
            kwScr=self.ltxtkw_script.text().strip().replace('*','%')
            if  self.cboTable.currentText()=='治具路徑校正值':
                colN=['治具路徑名稱', '條件1數值', '條件1單位', '條件2數值', '條件2單位', '校正數值', '校正單位']
                sSql=f"select  PName, Cond_1, Cond_1_unit, Cond_2, Cond_2_unit, CalValue, CalValue_unit from FixtureCal"                
            elif self.cboTable.currentText()=='機種測試腳本':
                colN=['機種型號', '測試腳本']
                sSql=f"select  Model, Script from ScriptRef"
                if kwScr!='':
                    sSql+=f" where Script like '{kwScr}'"
            elif self.cboTable.currentText()=='測試腳本':
                colN=['腳本名稱', '腳本版本', 'No_1', 'No_2', 'No_3', '測項名稱', '次項名稱', '次次項名稱'
                    , '條件1名稱', '條件1值', '條件1單位', '條件2名稱', '條件2值', '條件2單位', '條件3名稱', '條件3值', '條件3單位'
                    , '期望值', '單位', '下限值', '上限值', '偏差期望值', '偏差下限', '偏差上限'
                    , '條件_DUT', '條件_設備', '條件_治具', '條件_治具_其他', '備註']
                sSql=f"""select  Name, Version, ItemNo, SubItemNo, TestPointNo, ItemName, ItemName_2t, ItemName_3t 
                    , CondName_1, Cond_1, CondUnit_1, CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3 
                    , Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit  
                    , Cond_DUT, Cond_Equip, Cond_Fixture, Cond_Fixture_Other, Comment 
                    from TestScript"""
                if kwScr!='':
                    sSql+=f" where Name like '{kwScr}'"
            elif self.cboTable.currentText()=='治具路徑控制表':
                colN=['治具路徑名稱', '路徑描述', 'Port0腳位', 'Port1腳位', 'Port2腳位', 'Port3腳位', 'Port4腳位', 'Port5腳位']
                sSql=f"select  PName, PDesc, Port0, Port1, Port2, Port3, Port4, Port5 from FixturePath"
            elif self.cboTable.currentText()=='治具埠腳位':
                colN=['埠編號', '腳位']
                sSql=f"select PortNo, Pins from PortAssign_Mega"
            elif self.cboTable.currentText()=='測試腳本重號檢查':
                colN=['腳本名稱', 'No_1', 'No_2', 'No_3', '筆數']
                sSql=f"select name,Itemno,subItemNo,TestPointNo , count(name) as Qty from testscript group by name,Itemno,subItemNo,TestPointNo having Qty>=2"
            rsFix=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            cols=len(rsFix[0])
            rows=len(rsFix)
            self.table1.setColumnCount(cols)
            self.table1.setHorizontalHeaderLabels(colN) #設定欄名稱 
            for r in range(0,rows):
                self.table1.insertRow(r)
                for c in range(0, cols):
                    d=str(rsFix[r][c])
                    if (d is not None) and d!='None':
                        item=QTableWidgetItem(d)
                        self.table1.setItem(r, c, item)
        except Exception as e:
            print(f"{inspect.currentframe().f_code.co_name} \n Err:{str(e)}")


    def table_update(self):
        subProcName='table_update'
        # self.updateFixturePath()
        if self.ltxtkw_script.text()!='':
            sc=f"\n,且腳本條件為{self.ltxtkw_script.text()}的"
        else:
            sc=""
        reply = QMessageBox.information(self, '更新資料', f"將刪除資料庫{self.cboTable.currentText()}{sc}資料表內容\n再上傳畫面資料到資料庫.\n請再確認是否更新資料庫資料?", 
                            QMessageBox.No | QMessageBox.Yes, QMessageBox.No)
        if reply == QMessageBox.No:
            return
        d=[]
        try:            
            kwScr=self.ltxtkw_script.text().strip().replace('*','%')
            if  self.cboTable.currentText()=='治具路徑校正值':
                colN=['治具路徑名稱', '條件1數值', '條件1單位', '條件2數值', '條件2單位', '校正數值', '校正單位']
                TBN="FixtureCal" 
                sSql_d=f"delete from {TBN}"
                sSql=f"Insert into {TBN} (PName, Cond_1, Cond_1_unit, Cond_2, Cond_2_unit, CalValue, CalValue_unit) VALUES ("
            elif self.cboTable.currentText()=='機種測試腳本':
                colN=['機種型號', '測試腳本']
                TBN="ScriptRef"
                sSql_d=f"delete from {TBN}"
                if kwScr!='':
                        sSql_d+=f" where Script like '{kwScr}'"
                sSql=f"Insert into {TBN} (Model, Script) VALUES ("
            elif self.cboTable.currentText()=='測試腳本':
                colN=['腳本名稱','腳本版本' ,'No_1', 'No_2', 'No_3', '測項名稱', '次項名稱', '次次項名稱'
                    , '條件1名稱', '條件1值', '條件1單位', '條件2名稱', '條件2值', '條件2單位', '條件3名稱', '條件3值', '條件3單位'
                    , '期望值', '單位', '下限值', '上限值', '偏差期望值', '偏差下限', '偏差上限'
                    , '條件_DUT', '條件_設備', '條件_治具', '條件_治具_其他', '備註']
                TBN="TestScript"
                sSql_d=f"delete from {TBN}"
                if kwScr!='':
                        sSql_d+=f" where Name like '{kwScr}'"
                sSql=f"""Insert into {TBN} (Name, Version, ItemNo, SubItemNo, TestPointNo, ItemName, ItemName_2t, ItemName_3t 
                    , CondName_1, Cond_1, CondUnit_1, CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3 
                    , Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit  
                    , Cond_DUT, Cond_Equip, Cond_Fixture, Cond_Fixture_Other, Comment ) VALUES ("""
            elif self.cboTable.currentText()=='治具路徑控制表':
                colN=['治具路徑名稱', '路徑描述', 'Port0腳位', 'Port1腳位', 'Port2腳位', 'Port3腳位', 'Port4腳位', 'Port5腳位']
                TBN="FixturePath"
                sSql_d=f"delete from {TBN}"
                sSql=f"Insert into {TBN} (PName, PDesc, Port0, Port1, Port2, Port3, Port4, Port5) VALUES ("
            elif self.cboTable.currentText()=='治具埠腳位':
                colN=['埠編號', '腳位']
                TBN="PortAssign_Mega"
                sSql_d=f"delete from {TBN}"
                sSql=f"Insert into {TBN} (PortNo, Pins) VALUES ("
            for i in range(self.table1.columnCount()):
                if self.table1.horizontalHeaderItem(i).text()!=colN[i]:
                    QMessageBox.critical(self,'錯誤','資料欄位與要維護的資料表不符, \n無法更新資料!')
                    return
            sSql_d=sSql_d.replace('\0','')
            AteDB.ConnectSQLite(self.DBPth)
            AteDB.con.execute(sSql_d)
            for r in range(0,self.table1.rowCount()):
                d.clear()
                sSql_1=''
                bk=0
                for j in range(0,self.table1.columnCount()):
                    d.append(self.table1.item(r,j))
                    if d[j]!=None and (d[j] is not None):
                        ss= d[j].text().strip()
                        if ss!='': 
                            sSql_1+=f"'{ss}'"
                        else: 
                            sSql_1+=f"NULL"
                            bk+=1
                    else: 
                        sSql_1+=f"NULL"
                        bk+=1
                    if j<self.table1.columnCount()-1: sSql_1+=f", "
                    else: sSql_1+=f") "
                print(sSql+sSql_1)
                if bk<self.table1.columnCount():
                    sSql=sSql.replace('\0','')
                    sSql_1=sSql_1.replace('\0','')
                    AteDB.con.execute(sSql+sSql_1)
                else:
                    print(f"{j}:空白列, 不寫入資料庫")
            AteDB.con.commit() 
        except Exception as e:            
            QMessageBox.critical(self,'table_update 錯誤',f"{subProcName} Error: \n{str(e)}")
            AteDB.con.rollback()

        AteDB.CloseConnection() 

    #insert
    def table_insert(self):
        row = self.table1.rowCount()
        self.table1.insertRow(row)

    #delete
    def table_delete(self):
        row_select=self.table1.selectedRanges()
        if len(row_select) == 0:
            return
        print(f"len(row_select)=>{len(row_select)}")
        print(f"{row_select[0].topRow()},{row_select[0].bottomRow()}")
        i=0
        for row in range(row_select[0].topRow(),row_select[0].bottomRow()+1):
            print(f"removeRow({row})")
            self.table1.removeRow(row-i)        #列的位置會因, 刪除後改變, 需重計算
            i+=1 
        
    #從CSV檔匯入資料到表格    
    def table_import(self):   
        subProcName='table_import'
        colN=[]
        try:
            if self.FPth=='': self.FPth=path.dirname(self.DBPth)
            File_Name, File_Type=QFileDialog.getOpenFileName(self, "載入檔案", self.FPth, "csv Files (*.csv)")
            self.FPth=path.dirname(File_Name)
            print(File_Name, File_Type)
            kwScr=self.ltxtkw_script.text().strip().replace('*','%')
        
            if  self.cboTable.currentText()=='治具路徑校正值':
                colN=['治具路徑名稱', '條件1數值', '條件1單位', '條件2數值', '條件2單位', '校正數值', '校正單位']
            elif self.cboTable.currentText()=='機種測試腳本':
                colN=['機種型號', '測試腳本']
            elif self.cboTable.currentText()=='測試腳本':
                colN=['腳本名稱', '腳本版本','No_1', 'No_2', 'No_3', '測項名稱', '次項名稱', '次次項名稱'
                    , '條件1名稱', '條件1值', '條件1單位', '條件2名稱', '條件2值', '條件2單位', '條件3名稱', '條件3值', '條件3單位'
                    , '期望值', '單位', '下限值', '上限值', '偏差期望值', '偏差下限', '偏差上限'
                    , '條件_DUT', '條件_設備', '條件_治具', '條件_治具_其他', '備註']
            elif self.cboTable.currentText()=='治具路徑控制表':
                colN=['治具路徑名稱', '路徑描述', 'Port0腳位', 'Port1腳位', 'Port2腳位', 'Port3腳位', 'Port4腳位', 'Port5腳位']
            elif self.cboTable.currentText()=='治具埠腳位':
                colN=['埠編號', '腳位']
            
            if QFile.exists(File_Name):
                with open(File_Name, 'r') as stream:
                    self.table1.setRowCount(0)
                    self.table1.setColumnCount(0)
                    r=0
                    for rowdata in csv.reader(stream):
                        print(rowdata)
                        if r==0:
                            if len(rowdata)==len(colN):
                                self.table1.setColumnCount(len(rowdata))
                                self.table1.setHorizontalHeaderLabels(rowdata) #设置行表头 
                            else:
                                QMessageBox.critical(self, '錯誤', f"載入資料的欄位數不符!")
                                return
                        else:
                            row = self.table1.rowCount()
                            self.table1.insertRow(row)
                            for column, data in enumerate(rowdata):     #把rowdata分成索引及值二部分
                                item = QTableWidgetItem(data)
                                self.table1.setItem(row, column, item)
                        r+=1
                self.settings.setValue('defaultFilePath',self.FPth)
            else:
                QMessageBox.critical(self,"錯誤",'檔錯不存在!')
        except Exception as e:
            QMessageBox.critical(self,"錯誤",f"{subProcName} Error: \n{str(e)}")
    #匯出表格資料到CSV檔
    def table_export(self):
        try:
            if self.FPth=='': self.FPth=path.dirname(self.DBPth)
            File_Name, File_Type=QFileDialog.getSaveFileName(self, "匯出檔案", self.FPth, "csv Files (*.csv)")
            self.FPth=path.dirname(File_Name)
            print(File_Name, File_Type)
            if File_Name!='':
                rowdata = []
                with open(File_Name, 'w',newline='',encoding='big5') as stream:
                    writer = csv.writer(stream)
                    for c in range(self.table1.columnCount()):
                        item = self.table1.horizontalHeaderItem(c)
                        if item is not None:
                            rowdata.append(item.text())
                        else:
                            rowdata.append('')
                    writer.writerow(rowdata)
                    for row in range(self.table1.rowCount()):
                        rowdata.clear()
                        for column in range(self.table1.columnCount()):
                            item = self.table1.item(row, column)
                            if item is not None:
                                rowdata.append(item.text())
                            else:
                                rowdata.append('')
                        writer.writerow(rowdata)
            self.settings.setValue('defaultFilePath',self.FPth)
        except Exception as e:
            QMessageBox.critical(self,'table_export 錯誤', f"{str(e)} \n str({rowdata})")
    
    

    def oncboUpScriptChanged(self):
        QApplication.processEvents()
        try:            
            self.DBPth=window.DBPth
            AteDB.ConnectSQLite(self.DBPth)

            sSql=f"select Name, max(Version) as Version from Testscript where Name='{self.cboUpScript.currentText()}' group by Name"
            rsFix=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            cols=len(rsFix[0])
            rows=len(rsFix)
            if rows>0:
                self.ltxtUpVersion.setText(rsFix[0][1])
        except Exception as e:
            print(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}")

    def FillScriptSelector(self):
        try:            
            self.DBPth=window.DBPth
            AteDB.ConnectSQLite(self.DBPth)

            sSql=f"select distinct Name from Testscript"
            rsFix=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            cols=len(rsFix[0])
            rows=len(rsFix)
            self.cboUpScript.clear()
            if rows>0:
                for i in range(0,rows):
                    self.cboUpScript.addItem(rsFix[i][0])
        
        except Exception as e:
            self.txtMsg.append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}")
    
    #更新雲端脚本按鈕被按
    @pyqtSlot()
    def on_btnUploadToDB_clicked(self):
        QApplication.processEvents()
        self.Upload_to_DB()

    #上傳測試脚本到雲端資料庫
    def Upload_to_DB(self):
        QApplication.processEvents()
        try:            
            self.MariaDB_Host="172.16.1.209"
            self.MariaDB_port=3306
            # self.MariaDB_Name="ate"
            self.MariaDB_Name="ate_test"
            
            self.DBPth=window.DBPth
            AteDB.ConnectSQLite(self.DBPth)

            # sSql=f"""select Name, max(Version) as Version from Testscript where Name='{self.cboUpScript.currentText()}' group by Name"""
            sSql=f"""
                select Name, Version, ItemNo, SubItemNo, TestPointNo, ItemName, ItemName_2t, ItemName_3t,
                CondName_1, Cond_1, CondUnit_1,  CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3,
                Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit, 
                Cond_DUT, Cond_Equip, Cond_Fixture, Cond_Fixture_Other, Comment 
                from TestScript Where Name ='{self.cboUpScript.currentText()}' and Version ='{self.ltxtUpVersion.text()}'
                """ 
            print(sSql)
            rsScript=AteDB.Query(sSql)
            rows=len(rsScript)
            if rows > 0:
                self.txtMsg.setText(f"有{rows}筆資料")
                AteDB.CloseConnection()

                conn = pymysql.connect(
                    user=self.ltxtUserName.text(),
                    password=self.ltxtPassword.text(),
                    host=self.MariaDB_Host,
                    port=self.MariaDB_port,
                    database=self.MariaDB_Name)

                cur1 = conn.cursor()

                insert_sql = """
                    INSERT INTO TestScript (Name, Version, ItemNo, SubItemNo, TestPointNo, ItemName, ItemName_2t, ItemName_3t,
                    CondName_1, Cond_1, CondUnit_1, CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3,
                    Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit, 
                    Cond_DUT, Cond_Equip, Cond_Fixture, Cond_Fixture_Other, Comment)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """

                for row in rsScript:
                    cur1.execute(insert_sql, tuple(row))

                conn.commit()
                conn.close()
            else:
                self.txtMsg_append("查無資料，未執行上傳", QColorConstants.Red)
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",QColorConstants.Red)

    #附加文字到txtMsg
    def txtMsg_append(self,s="",TextColor=QColor("black")):        
        self.txtMsg.setTextColor(TextColor)
        self.txtMsg.append(s)
        self.txtMsg.verticalScrollBar().setValue(self.txtMsg.verticalScrollBar().maximum()) #垂宜捲到最下
        QtWidgets.QApplication.processEvents()
         


class MainTest(QtWidgets.QMdiSubWindow, Ui_frmMainTest):#UI 測試項目請在這修改and Data欄位Item Name也須修改一樣   #UI 治具條件列表請修改Data欄位Cond_Fixture     
    StopTest=False
    BRList=[]       #使用的Baud Rate List
    SkipPort=[]     #使用大寫儲存不檢查的連接埤    
    #已程式完成項目, 讓測項自動預設勾選----------------------------------------------------------------------------------
    OKItem=['PEL_CC準確度','PEL_CC表頭準確度','PEL_CR準確度','PEL_CV準確度','PEL_CV表頭準確度','PEL_CP準確度','電壓設定準確度','空戴漏電流確認', '空載漏電流確認', 'IR電壓設定準確度', 'IR電壓調整率', '電壓調整率'
            ,'電壓量測準確度', '電流量測準確度','Cut off 電流準確度', 'IR絕緣阻抗量測準確度','GB電流設定準確度'
            ,'GB電阻量測準確度','寫入機器序號','背板輸出確認', '掃描功能確認', 'ARC功能確認', 'REMOTE功能確認'
            ,'Signal IO功能確認', 'GPIB功能確認','寫入USB識別碼','寫入USB識別碼確認','Initial','CONT電阻量測準確度','寫入機器日期時間']  
    #------------------------------------------------------------------------------------------------------------------     
    SpecTRow=[]     #儲存指定要測試的列
    SpecTRowFix=[]     #儲存指定要測試的列,所用治具
    default_ACW_HiSet=40
    default_DCW_HiSet=8
    default_GB_HiSet=650
    #類別初始程序
    def __init__(self, parent=None):
        super(MainTest,self).__init__(parent)        
        self.setupUi(self) 
        
        self.DBName='ATE_PEL3000.db'
        self.MyDTFormat='yyyy-MM-dd hh:mm:ss'       #日期時間格式, 作為後面取得日期時間字串使用
        self.MyTimeFormart='hh:mm:ss.ss'

        # 建立 QSqlDatabase 資料庫物件 
        self.MariaDB = QSqlDatabase.addDatabase('QMYSQL')
        self.MariaDB_User="g0255"
        self.MariaDB_Password="PE_Go_wc9999"
        self.MariaDB_Host="172.16.1.209"
        self.MariaDB_port=3306
        self.MariaDB_Name="ate_test"
        self.MariaDB.setHostName(self.MariaDB_Host)
        self.MariaDB.setPort(self.MariaDB_port)
        self.MariaDB.setDatabaseName(self.MariaDB_Name)
        self.MariaDB.setUserName(self.MariaDB_User)
        self.MariaDB.setPassword(self.MariaDB_Password)

        self.EquipsBeUsed=''    #記錄那些設備連線使用
        self.settings = QSettings("txtSkipPort")
        self.SwitchPath=''      #目前治具開關路徑
        self.checkedFix=[]
        self.FirstExecPQC=True
        self.DUT=DUT_USB

        self.GBVLimit=5.4   #GB最大電壓值, GB電流*GB Hiset
        self.DUTIDN=''      #DUT的IDN資訊, updateTestSummary會用到此變數 
        
        font = self.txtSpecTRow.font()
        font.setBold(True)
        self.txtSpecTRow.setFont(font)
        self.txtSpecTRow.setTextColor(color_Magenta)        
        try:
            self.db =  QSqlDatabase.addDatabase('QSQLITE')      #建立資料庫物件
            self.txtSkipPort.setText(self.settings.value("txtSkipPort"))
            df=f'{str(pathlib.Path(__file__).parent.absolute())}\\{self.DBName}'
            if path.isfile(self.settings.value("DBPath")) and fnmatch(self.settings.value("DBPath"),'*.db') and access(self.settings.value("DBPath"),os.W_OK):
                self.ledDBPath.setText(self.settings.value("DBPath"))
            elif path.isfile(df):
                self.ledDBPath.setText(df)
            else:
                self.ledDBPath.setText('')
            self.DBPth=self.ledDBPath.text()
            self.ltxtInspector.setText(self.settings.value("Inspector"))           
            self.chkFixOptimize.setChecked(bool(self.settings.value("chkFixOptimize")) )
            self.chkFixManual.setChecked(bool(self.settings.value("chkFixManual")))
            self.chkHVMManual.setChecked(bool(self.settings.value("chkHVMManual")))
            if self.settings.value("ltxtDUTModel")!='':    
                self.ltxtDUTModel.setText(self.settings.value("ltxtDUTModel")) 
            self.ltxtDUTSn.setText(self.settings.value("ltxtDUTSn"))                  
        except Exception as e:
            QMessageBox.critical(self,"MainTest __init__ 錯誤",str(e))            
        
        try:
            self.tbView.setGeometry(QtCore.QRect(0, 0, self.scrollArea_2.width(), self.scrollArea_2.height()))
            self.stimer=0       #開始計時器
            self.gboxTestItems=MyQGroupBox()
            self.gboxFix=MyQGroupBox()
            self.GLayout=QGridLayout()
            self.GLayout_F=QGridLayout()
            self.gboxTestItems.setLayout(self.GLayout)
            self.scrollArea.setWidget(self.gboxTestItems)
            self.gboxTestItems.setTitle('測試項目')
            self.gboxFix.setLayout(self.GLayout_F)
            self.scrollArea_3.setWidget(self.gboxFix)
            self.gboxFix.setTitle('治具條件列表')
            self.cboScript.setEditable(True)
            self.getTestScript()
            self.selectScript()
            self.showPathItem()
            self.showPathFix()
            self.showEquip()            
            self.setDevInfo()            

            self.timer1=QtCore.QTimer()
            self.timer1.timeout.connect(self.onTimerTick)
            self.btnExit.clicked.connect(self.onExitClick)
            self.btnStart.clicked.connect(self.onStartClick)
            self.btnStop.clicked.connect(self.onStopClick)
            self.btnNewReport.clicked.connect(self.onNewReport)
            self.btnConDev.clicked.connect(self.on_btnConDev_Click)
            self.btnProTest.clicked.connect(self.on_btnProTest_Click)
            self.btnFixtureTest.clicked.connect(self.on_btnFixtureTest_Click)
            self.btnMModify.clicked.connect(self.on_btnMModify_Click)
            self.cboScript.currentIndexChanged.connect(self.oncboScriptChanged)
            self.cboReportNo.currentIndexChanged.connect(self.oncboReportNoChanged)
            self.ledDBPath.textChanged.connect(self.on_ledDBPath_textChanged)
            self.txtSpecTRow.textChanged.connect(self.on_txtSpecTRow_textChanged)
            self.btnExportRec.clicked.connect(self.on_btnExportRec_Click)

            for i in range(0,self.model.rowCount()):  
                rdev=self.model.record(i) 
                rdev.setValue('Connected',0)
                self.model.setRecord(i,rdev)
                self.model.submitAll()
            self.showEquip()            
            self.onlyStopDisable()  #把stop先設為不能按

            self.default_Width=self.width()
            self.default_Heigh=self.height()
            self.default_txtMsg_W=self.txtMsg.width()
            self.default_txtMsg_H=self.txtMsg.height()
            self.default_scrollArea_W=self.scrollArea.width()
            self.default_scrollArea_H=self.scrollArea.height()
            self.default_scrollArea_2_W=self.scrollArea_2.width()
            self.default_scrollArea_2_H=self.scrollArea_2.height()
            self.default_scrollArea_3_W=self.scrollArea_3.width()
            self.default_scrollArea_3_H=self.scrollArea_3.height()
            self.default_scrollArea_3_y=self.scrollArea_3.geometry().y()
            self.default_buttom_x=self.btnStart.geometry().x()

            self.setTabOrder(self.ltxtInspector,self.ltxtDUTModel)
            self.setTabOrder(self.ltxtDUTModel,self.ltxtDUTSn)
            self.setTabOrder(self.ltxtDUTSn,self.btnStart)
            self.setTabOrder(self.btnStart,self.btnNewReport)
            self.setTabOrder(self.btnNewReport,self.btnExit)
            self.setTabOrder(self.btnExit, self.btnConDev)
            self.setTabOrder(self.btnConDev,self.txtSpecTRow)
            self.setTabOrder(self.txtSpecTRow,self.btnMModify)
                        
            QtWidgets.QApplication.processEvents()
        except Exception as e:            
            QMessageBox.critical(self,"MainTest __init__ 錯誤",str(e))

    #啟動後, 還要執行的程式
    def appStarted(self):
        # 獲取主機名
        self.MyHostName = self.get_hostname()
        # #獲取IP
        # MyIp = self.get_ip()

        self.ltxtWorkstation.setText(self.MyHostName)        

        self.checkDBField() #檢查db檔內容
        self.chekTable_TestLog()    #檢查是否有TestLog資料表
    
    #取得主機名稱
    def get_hostname(self):
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
    def get_ip(self):
        try:
            ip = subprocess.check_output(['hostname', '-I']).decode('utf-8').strip()
            return ip
        except subprocess.CalledProcessError:
            return "Unknown IP"
    
    #變動視窗大小處理程序    
    def resizeEvent(self, event): 
        super().resizeEvent(event)
        try:       
            width, height = event.size().width(), event.size().height()
            if self.isMaximized():  #不要讓視窗最大化
                self.showNormal()
                
            if not self.isMinimized():
                MinH=self.default_Heigh-160
                if height>self.default_Heigh:
                    h=self.default_Heigh
                elif height<MinH:
                    h=MinH
                else:
                    h=height
                w=width
                if width>1200: w=1200
                if width<700: w=700
                self.resize(w,h)
                dw=w-self.default_Width
                dh=h-self.default_Heigh
                self.txtMsg.resize(self.default_txtMsg_W+dw,self.default_txtMsg_H)
                self.ledDataUpload.resize(self.txtMsg.width(), self.ledDBPath.height())
                
                self.scrollArea_2.resize(self.default_scrollArea_2_W+dw,self.default_scrollArea_2_H+int(dh/2))
                self.tbView.resize(self.default_scrollArea_2_W+dw,self.default_scrollArea_2_H+int(dh/2))
                self.scrollArea_3.resize(self.default_scrollArea_3_W+dw,self.default_scrollArea_3_H+int(dh/2))
                self.scrollArea_3.move(self.scrollArea_3.x(),self.default_scrollArea_3_y+int(dh/2))
                self.scrollArea.resize(self.default_scrollArea_W,self.default_scrollArea_H+dh)
                
                #移動控制項
                try:            
                    btns=self.findChildren(QPushButton)
                    for bt in btns:
                        bt.move(self.default_buttom_x+dw,bt.y())
                    self.lblSkipPort.move(self.default_buttom_x+dw,self.lblSkipPort.y())
                    self.lblSpecTRow.move(self.default_buttom_x+dw,self.lblSpecTRow.y())
                    self.txtSkipPort.move(self.default_buttom_x+dw,self.txtSkipPort.y())
                    self.txtSpecTRow.move(self.default_buttom_x+dw,self.txtSpecTRow.y())
                    self.chkFixManual.move(self.default_buttom_x+dw,self.chkFixManual.y())
                    self.chkFixOptimize.move(self.default_buttom_x+dw,self.chkFixOptimize.y())
                    self.chkHVMManual.move(self.default_buttom_x+dw,self.chkHVMManual.y())
                except Exception as e: QMessageBox.critical(self,"MainTest resizeEvent 錯誤",str(e))
        except Exception as e: QMessageBox.critical(self,"MainTest resizeEvent 錯誤",str(e))
    
    #產生新的測試編號
    def getTestNo(self):
        sn=self.ltxtDUTSn.text()

        tno=''
        try:
            if sn!='':
                s='yyyyMMddhhmmss'
                datetime = QDateTime.currentDateTime()        
                tno=f"{sn}_{datetime.toString(s)}"
            else:
                self.txtMsg_append(f'無序號,無法取得測試編號!', Qt.red)
                tno=''
        except Exception as e: 
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: {str(e)}", Qt.red) 
        return tno

    #要產生新測試資料, 在資料表建立相關測試點的資料
    def onNewReport(self):
        TNo=self.getTestNo()        
        try:
            if TNo !='' and self.cboScript.currentText()!='' and self.DBPth!='':                
                AteDB.ConnectSQLite(self.DBPth)
                sSql=f"select Model, script from ScriptRef where Model='{self.ltxtDUTModel.text()}'"
                rsScr=AteDB.Query(sSql)
                if len(rsScr)>0:
                    ScriptOK=False
                    for i in range(0,len(rsScr)):
                        if rsScr[i][1]==self.cboScript.currentText():
                            ScriptOK=True
                            break
                    if ScriptOK==False:
                        self.txtMsg_append(f"型號與測試腳本不符,無法建立新測試文件!",Qt.red)
                    else:
                        sSql=f"select Name, Max(Version) as Version from TestScript group by Name having name='{self.cboScript.currentText()}'"
                        rsVer=AteDB.Query(sSql)
                        self.ltxtScriptVer.setText(rsVer[0][1])     #設定版本為最新版本
                        sSql=f'''insert into TestRecord (TestDocNo, ScriptName, ScriptVersion, ItemNo, SubItemNo, TestPointNo, ItemNo, ItemName, ItemName_2t, ItemName_3t, CondName_1, Cond_1, CondUnit_1, 
                                CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3, Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit, Cond_Fixture, Cond_DUT)
                                select '{TNo}', name, Version, ItemNo, SubItemNo, TestPointNo, ItemNo, ItemName, ItemName_2t, ItemName_3t, CondName_1, Cond_1, CondUnit_1, 
                                CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3,Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit, Cond_Fixture, Cond_DUT 
                                from TestScript where name = '{self.cboScript.currentText()}' and Version='{rsVer[0][1]}' '''
                        sSql=sSql.replace('\0','')
                        AteDB.con.execute(sSql)
                        AteDB.con.commit()
                        AteDB.CloseConnection()
                        self.cboReportNo.addItem(TNo)
                        self.cboReportNo.setCurrentIndex(self.cboReportNo.count()-1)
                        #self.showPathItem()
                        self.updateTestSummary()
                else:
                    self.txtMsg_append(f"查不到型號對應的測試腳本,無法建立新測試文件!",Qt.red)
            else:
                self.txtMsg_append(f'無序號,測試腳本或找不到資料檔,無法建立新測試文件!',Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
    
    #更新TestSummary(測試總表)資料表
    def updateTestSummary(self):
        QApplication.processEvents()
        TDocNo=self.cboReportNo.currentText()
        try:  
            AteDB.ConnectSQLite(self.DBPth)
            sSql=f"select * from TestSummary where TestDocNo='{TDocNo}'"            
            rsTS=AteDB.Query(sSql)
            s='yyyy/MM/dd hh:mm:ss'
            datetime = QDateTime.currentDateTime()        
            ldt=datetime.toString(s)
            #計算一份測試報告的測項數, Pass, Fail, 未測的項數
            sSql=f"""SELECT A.TestDocNo, A.TestResult, A.Workstation, A.Staff, B.PASS, B.FAIL, B.其他, B.小計 
                    FROM TestSummary A LEFT JOIN (
                        SELECT TestDocNo ,
                            SUM(1) '小計', 
                            SUM(CASE WHEN Judge = 'PASS' THEN 1 ELSE 0 END) 'PASS' ,
                            SUM(CASE WHEN Judge = 'FAIL' THEN 1 ELSE 0 END) 'FAIL' ,
                            SUM(CASE WHEN Judge <> 'PASS' AND Judge <> 'FAIL' THEN 1 ELSE 0 END) '其他'
                        FROM TestRecord GROUP BY TestDocNo ) B ON A.TestDocNo = B.TestDocNo
                    WHERE A.TestDocNo='{TDocNo}'"""
            rsTResult=AteDB.Query(sSql)
            TJF=0
            TJP=0
            TestJ='' 
            self.ltxtInspector.setText(self.ltxtInspector.text().upper().strip())
            QApplication.processEvents()
            if len(rsTS)==0:        
                sSql=f'''Insert into TestSummary (TestDocNo, ScriptName,  ScriptVersion, Workstation, CreateDatetime, LastTestDatetime, DUTModel, DUTSerNo, Staff, TestResult) 
                        VALUES ('{TDocNo}','{self.cboScript.currentText()}', '{self.ltxtScriptVer.text()}' ,'{self.ltxtWorkstation.text()}','{ldt}','{ldt}', '{self.ltxtDUTModel.text()}'
                        , '{self.ltxtDUTSn.text()}','{self.ltxtInspector.text()}' , '') '''
            else:
                sSql=f"""SELECT A.TestDocNo, A.TestResult, B.PASS, B.FAIL, B.其他, B.小計 
                    FROM TestSummary A LEFT JOIN (
                        SELECT TestDocNo ,
                            SUM(1) '小計', 
                            SUM(CASE WHEN Judge = 'PASS' THEN 1 ELSE 0 END) 'PASS' ,
                            SUM(CASE WHEN Judge = 'FAIL' THEN 1 ELSE 0 END) 'FAIL' ,
                            SUM(CASE WHEN Judge <> 'PASS' AND Judge <> 'FAIL' THEN 1 ELSE 0 END) '其他'
                        FROM TestRecord GROUP BY TestDocNo ) B ON A.TestDocNo = B.TestDocNo
                    WHERE A.TestDocNo='{TDocNo}'"""
                rsTResult=AteDB.Query(sSql)
                if rsTResult[0][2]==rsTResult[0][5]:
                    TestJ='PASS'
                elif rsTResult[0][3]>0:
                    TestJ='FAIL'
                else:
                    TestJ=''
                sSql=f'''Update TestSummary SET Workstation='{self.ltxtWorkstation.text()}', LastTestDatetime='{ldt}', TestResult='{TestJ}'
                        , DUTModel='{self.ltxtDUTModel.text()}', DUTIDN='{self.DUTIDN}', Staff='{self.ltxtInspector.text()}'
                        , EquipList='{self.EquipsBeUsed}', Datetime_Upload=NULL 
                        Where TestDocNo='{TDocNo}'  '''
            sSql=sSql.replace('\0','')
            AteDB.con.execute(sSql)
            AteDB.con.commit()
            AteDB.CloseConnection()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: {str(e)} \n SQL= {sSql}",Qt.red)
            print(f"sSql={sSql}")
            
    #取得測試腳本測項名稱
    def getTestItems(self):
        rsTItem=[]
        try:
            AteDB.ConnectSQLite(self.DBPth)
            TDocNo=self.cboReportNo.currentText()
            sSql=f'select distinct ItemNo, ItemName, ScriptName from TestRecord where TestDocNo=\'{TDocNo}\' order by ItemNo'            
            rsTItem=AteDB.Query(sSql)        
            AteDB.CloseConnection()
        except Exception as e:            
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        return rsTItem
    
    #取得治具條件列表
    def getFixCond(self):
        rsFix=[]
        rsT=[]
        try:            
            AteDB.ConnectSQLite(self.DBPth)
            TDocNo=self.cboReportNo.currentText()
            sSql=f'select distinct Cond_Fixture from TestRecord where TestDocNo=\'{TDocNo}\''
            rsFix=AteDB.Query(sSql)        
            AteDB.CloseConnection()            
            j=len(rsFix)-1
            while j>=0:
                if rsFix[j][0].find('手動')>=0:
                    rsT.append(rsFix[j])
                    del rsFix[j]
                j-=1
            for r in rsT:
                rsFix.append(r)                    
        except Exception as e:            
           self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        return rsFix

    #取得治具指定路徑及條件值的校證值
    def getFixCalValue(self,FixPath='', Cond_1=''):        
        CValue=''
        CVUnit=''
        try:
            AteDB.ConnectSQLite(self.DBPth)
            sSql="select PName, Cond_1, Cond_1_unit, CalValue, CalValue_unit from FixtureCal "
            if Cond_1!='' and (Cond_1 is not None):
                sSql+=f"where PName='{FixPath}' and Cond_1='{Cond_1}'"
            else:
                sSql+=f"where PName='{FixPath}'"
            rsCalV=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            if len(rsCalV)>0:
                CValue=rsCalV[0]['CalValue']
                CVUnit=rsCalV[0]['CalValue_unit']
            else:
                CValue=''
                CVUnit=''
        except Exception as e:            
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        return CValue, CVUnit
    
    #取得測試脚本的名稱
    def getTestScript(self):
        rsScript=[]
        try:
            self.cboScript.clear()
            AteDB.ConnectSQLite(self.DBPth)            
            sSql='select distinct Name from TestScript order by Name'              
            rsScript=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            for i in range(0,len(rsScript)): 
                self.cboScript.addItem(str(rsScript[i][0]))        
        except Exception as e:            
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        return rsScript
    
    #取得該序號己有的測試編號
    def getSNDocNo(self):
        rsDNo=[]
        SerNo=self.ltxtDUTSn.text()
        try:
            AteDB.ConnectSQLite(self.DBPth)            
            sSql=f"select distinct TestDocNo from TestRecord where TestDocNo like '{SerNo}_%' order by TestDocNo desc"   
            print(sSql)           
            rsDNo=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            self.cboReportNo.clear()
            if len(rsDNo)>0:
                for i in range(0,len(rsDNo)): 
                    self.cboReportNo.addItem(str(rsDNo[i][0])) 
            else:
                self.onNewReport()
        except Exception as e:            
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        return rsDNo
    #檢查是否有TestLog資料表, 若没有自動建立
    def chekTable_TestLog(self):
        try:
            if self.db.open():
                # 檢查資料表是否存在
                query = QSqlQuery()
                query.prepare(f"SELECT name FROM sqlite_master WHERE type='table' AND name='TestLog'")
                query.exec_()
                if not(query.next()):   # 資料表不存在
                    query.prepare(f"""    CREATE TABLE IF NOT EXISTS TestLog (
                                    TestDocNo	STRING,
                                    ItemNo	INT,
                                    SubItemNo	INT,
                                    TestPointNo	INT,
                                    Workstation	STRING,
                                    Staff	STRING,
                                    Measure_Main	DOUBLE,
                                    Measure_1	DOUBLE,
                                    Measure_2	DOUBLE,
                                    Judge	STRING,
                                    Comment	STRING,
                                    Datetime_Start	Datetime,
                                    Datetime_End	Datetime,
                                    Datetime_Upload	Datetime
                                )""")
                    query.exec_()
                    self.db.commit()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
    #檢查資料表是否有需要的欄位, 若没有, 自動加入
    def Table_Add_Colum(self, TableName='', ColumName='', DataType=''):
        try:
            if self.db.open():
                # 檢查資料表是否存在
                query = QSqlQuery()
                query.prepare(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{TableName}'")
                query.exec_()
                if query.next():
                    # 檢查欄位是否存在
                    query.prepare(f"PRAGMA table_info({TableName})")
                    query.exec_()
                    timestamp_exists = False
                    while query.next():
                        if query.value(1) == ColumName:
                            timestamp_exists = True
                            break
                    if timestamp_exists:
                        print(f"資料表 {TableName} 中已經存在 {ColumName} 欄位")
                    else:
                        # 新增欄位
                        query.prepare(f"ALTER TABLE {TableName} ADD COLUMN {ColumName} {DataType}")
                        query.exec_()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red) 
    #檢查資料表的欄位資料型別是否為TEXT若不是, 更改之
    def Table_Alter_Datatype(self, TableName='', ColumName=[], pk_col=[]):
        try:
            if self.db.open():
                strCol=", ".join(ColumName)
                # 檢查 TableName 欄位 ColumName 的資料類型
                query = QSqlQuery(self.db)
                query.exec_(f"SELECT {strCol} FROM {TableName} LIMIT 1")
                needAlter=False
                for r in range(query.record().count()):
                    print(query.record().field(r).type())
                    if query.record().field(r).type() !=QVariant.String:        # 取得資料型態, 並判斷是否變更
                        needAlter=True
                        break
                # 如果 ColumName 的資料型態不是 DataType，則執行修改
                if needAlter:                
                    fields, Name_List=self.get_table_fields(TableName)                          #取得資料欄位資訊
                    self.db.transaction()
                    query.exec_(f"DROP TABLE Table_backup;")
                    query.exec_(f"CREATE TABLE Table_backup AS SELECT * FROM {TableName}")     #建立臨時的資料表Table_backup，並複製原本的資料                
                    print(query.lastError().text())
                    
                    self.db.commit()
                    self.db.transaction()
                    query.exec_(f"DROP TABLE {TableName};")                                     #刪除原本資料表 
                    print(query.lastError().text())
                    strS1=''
                    strS2=''
                    strS3='' 
                    strPK=''
                    
                    strPK=", ".join(pk_col)              
                    for i in range(len(Name_List)):
                        if Name_List[i] in ColumName:
                            strS1=strS1+f", {Name_List[i]} TEXT"
                            strS3=strS3+f", CAST({Name_List[i]} AS TEXT)"
                        else:
                            strS1=strS1+f", {fields[i]}"
                            strS3=strS3+f", {Name_List[i]}"
                        strS2=strS2+f", {Name_List[i]}"                    
                    strS1=strS1[1:]                     #欄名 型別,
                    strS2=strS2[1:]                     #欄名,
                    strS3=strS3[1:]                     #欄名, 部份加入CAST
                    strSQL=f"CREATE TABLE {TableName} ({strS1}, CONSTRAINT {TableName}_pk PRIMARY KEY ({strPK}));"
                    # strSQL=f"CREATE TABLE {TableName} ({strS1}) ;"
                    print(strSQL )
                    query.exec_(strSQL)     #建立新的資料表
                    print(query.lastError().text())
                    strSQL=f"INSERT INTO {TableName} ({strS2}) SELECT {strS3} FROM Table_backup"
                    print(strSQL )
                    query.exec_(strSQL)     #匯入Table_backup資料
                    print(query.lastError().text())
                    # query.exec_("DROP TABLE Table_backup;")         #刪除臨時的資料表Table_backup 
                    strSQLErr=query.lastError().text()
                    if strSQLErr =="":               
                        self.db.commit()
                    else:
                        self.txtMsg_append(f"SQL Error: \n{str(strSQLErr)}",Qt.red) 
                        self.db.rollback()
                else:
                    print(f"{ColumName} 已經是 TEXT")   
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
            self.db.rollback()             
        QApplication.processEvents()    
    
    def get_table_fields( self,table_name):        
        try:
            if not self.db.open():
                return 'Unable to establish a database connection.'        
            # 查詢Table_A欄位資訊
            query = QSqlQuery(f"PRAGMA table_info({table_name})", self.db)
            Fields = []
            Field_Name_List=[]
            while query.next():
                field_name = query.value(1)
                field_type = query.value(2)
                Fields.append(f"{field_name} {field_type}")
                Field_Name_List.append(field_name)
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        # 回傳欄位資訊字串
        return Fields, Field_Name_List
    
    
    #檢查測試記錄表新增的欄及後續變更的資料型是否正確
    def checkDBField(self): 
        try:       
            if self.db.open():
                query=QSqlQuery(self.db)
                self.Table_Add_Colum("TestRecord","Datetime_Upload","Datetime")
                self.Table_Add_Colum("TestSummary","Datetime_Upload","Datetime")
                self.Table_Add_Colum("TestScript","Version", "Text")
                self.Table_Add_Colum("TestSummary","ScriptVersion", "Text")
                self.Table_Add_Colum("TestRecord","ScriptVersion", "Text")               
                
                query.prepare(f"Update TestScript set Version='' where Version IS NULL")      #把版本為NULL的, 改為空字串
                query.exec_()
                
                # self.Table_Alter_Datatype("TestRecord",["ItemNo","SubItemNo","TestPointNo"],["TestDocNo","ItemNo","SubItemNo","TestPointNo"])                
                # self.db.commit()
                # self.Table_Alter_Datatype("TestScript",["ItemNo","SubItemNo","TestPointNo"],["Name","ItemNo","SubItemNo","TestPointNo"]) 
                # self.db.commit()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)         
        QApplication.processEvents()

    #上傳測試資料
    def UploadTestData(self): 
        #--------------測試時在開發電腦不上傳-------------
        if self.MyHostName.lower().find('lorenzo')>=0:
            return
        #---------------------------------
        try:            
            conn = pymysql.connect(
                user=self.MariaDB_User,
                password=self.MariaDB_Password,
                host=self.MariaDB_Host,
                port=self.MariaDB_port,
                database=self.MariaDB_Name)
            
            # self.MariaDB.open()
            query = QSqlQuery(self.db) 
            # query2 = QSqlQuery(self.MariaDB)
            # query3 = QSqlQuery(self.MariaDB)
            query4 = QSqlQuery(self.db)

            cur1=conn.cursor()
            cur2=conn.cursor()

            #===上傳TestRecord資料表===            
            query.prepare('''SELECT TDatetime, TestDocNo, ScriptName, ScriptVersion ,ItemNo, SubItemNo, TestPointNo, ItemName, 
                            ItemName_2t, ItemName_3t, CondName_1, Cond_1, CondUnit_1, CondName_2, Cond_2, CondUnit_2, 
                            CondName_3, Cond_3, CondUnit_3, Expect, ExpectUnit, LowLimit, UpLimit, Cond_DUT, Cond_Equip, 
                            Measure_Main, Measure_1, Measure_2, Deviation, Deviation_Expect, Deviation_LowLimit, 
                            Deviation_UpLimit, Judge, Comment, Cond_Fixture, Cond_Fiixture_Other, Datetime_Upload 
                            FROM TestRecord WHERE Datetime_Upload IS NULL            
                            ''')
            query.exec_()
            if query.lastError().text() !='':
                print(f"{inspect.currentframe().f_code.co_name} SQL錯誤訊息:\n{query.lastError().text()}")
            
            #檢查並更新或新增到 MariaDB 資料表
            c=[]
            self.txtMsg_append(f"上傳測試記錄中...")   
            urow=0         
            while query.next():
                if self.StopTest: break
                urow+=1
                self.ledDataUpload.setText(f"上傳記錄第{urow}筆")
                QApplication.processEvents()
                c.clear()
                for i in range(query.record().count()):
                    if i in [0,19,21,22,25,26,27,28,29,30,31] and query.value(i)=='':
                        c.append('NULL')
                    else:
                        c.append(query.value(i))
                # 檢查 MariaDB 資料表 是否已經存在相同的記錄   
                strSQL=(f"SELECT TestDocNo, ItemNo, SubItemNo, TestPointNo " + 
                        f"FROM TestRecord " + 
                        f"WHERE TestDocNo = '{c[1]}' AND ItemNo = {c[4]} AND SubItemNo = {c[5]} AND TestPointNo = {c[6]}")
                
                # print(strSQL)
                cur1.execute(strSQL)
                result1=cur1.fetchall()
                rcount=len(result1)
                
                UpDateTime=QDateTime.currentDateTime().toString(self.MyDTFormat)
                # self.MariaDB.transaction()                
                if rcount>0:
                    # 記錄已存在，使用更新方式
                    strSQL=(f"UPDATE TestRecord SET TDatetime = '{c[0]}', " +  
                            f"TestDocNo = '{c[1]}', ScriptName = '{c[2]}', ScriptVersion = '{c[3]}', ItemNo = {c[4]}, SubItemNo = {c[5]}, TestPointNo = {c[6]}, " + 
                            f"ItemName = '{c[7]}', ItemName_2t = '{c[8]}', ItemName_3t = '{c[9]}', CondName_1 = '{c[10]}', Cond_1 = '{c[11]}', " +
                            f"CondUnit_1 = '{c[12]}', CondName_2 = '{c[13]}', Cond_2 = '{c[14]}', CondUnit_2 = '{c[15]}', CondName_3 = '{c[16]}', " +
                            f"    Cond_3 = '{c[17]}', CondUnit_3 = '{c[18]}', Expect = {c[19]}, ExpectUnit = '{c[20]}', LowLimit = {c[21]}, " +
                            f"    UpLimit = {c[22]}, Cond_DUT = '{c[23]}', Cond_Equip = '{c[24]}', Measure_Main = {c[25]}, Measure_1 = {c[26]},  " +
                            f"    Measure_2 = {c[27]}, Deviation = {c[28]}, Deviation_Expect = {c[29]}, Deviation_LowLimit = {c[30]}, Deviation_UpLimit = {c[31]},  " +
                            f"    Judge = '{c[32]}', Comment = '{c[33]}', Cond_Fixture = '{c[34]}', Cond_Fiixture_Other = '{c[35]}', Datetime_Upload = '{UpDateTime}' " +
                            f"    WHERE TestDocNo = '{c[1]}' AND ItemNo = {c[4]} AND SubItemNo = {c[5]} AND TestPointNo = {c[6]} " )
                    strSQL=strSQL.replace("'NULL'","NULL")                    
                    # print(strSQL)
                    cur2.execute(strSQL)
                else:
                    # 記錄不存在，使用新增方式
                    strSQL=f"""INSERT INTO TestRecord (TDatetime, TestDocNo, ScriptName, ScriptVersion, ItemNo, SubItemNo, TestPointNo, ItemName, 
                                ItemName_2t, ItemName_3t, CondName_1, Cond_1, CondUnit_1, CondName_2, Cond_2, CondUnit_2, 
                                CondName_3, Cond_3, CondUnit_3, Expect, ExpectUnit, LowLimit, UpLimit, Cond_DUT, Cond_Equip, 
                                Measure_Main, Measure_1, Measure_2, Deviation, Deviation_Expect, Deviation_LowLimit, 
                                Deviation_UpLimit, Judge, Comment, Cond_Fixture, Cond_Fiixture_Other, Datetime_Upload) 
                                VALUES ('{c[0]}', '{c[1]}', '{c[2]}', '{c[3]}', {c[4]}, {c[5]}, {c[6]}, '{c[7]}', '{c[8]}', '{c[9]}', '{c[10]}', 
                                '{c[11]}', '{c[12]}', '{c[13]}', '{c[14]}', '{c[15]}', '{c[16]}', '{c[17]}', '{c[18]}', {c[19]}, '{c[20]}', 
                                {c[21]}, {c[22]}, '{c[23]}', '{c[24]}', {c[25]}, {c[26]}, {c[27]}, {c[28]}, {c[29]}, {c[30]},
                                {c[31]}, '{c[32]}', '{c[33]}', '{c[34]}', '{c[35]}', '{UpDateTime}')"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    cur2.execute(strSQL)                
                conn.commit()
                strSQLErr=query.lastError().text()                
                if strSQLErr =="":
                    query4.prepare(f"""UPDATE TestRecord SET Datetime_Upload = :c0 
                                WHERE TestDocNo = '{c[1]}' AND ItemNo = {c[4]} AND SubItemNo = {c[5]} AND TestPointNo = {c[6]}"""
                                )
                    query4.bindValue(':c0', UpDateTime)
                    query4.exec_()
                    self.db.commit()
                else:
                    self.txtMsg_append(f"SQL Error: \n{str(strSQLErr)}",Qt.red) 

            #===上傳TestSummary資料表===
            # query = QSqlQuery(self.db)
            query.prepare('''SELECT TestDocNo, ScriptName, ScriptVersion, Workstation, CreateDatetime, LastTestDatetime, TestResult, 
                            DUTModel, DUTSerNo, DUTIDN, Temperature, Humidity, EquipList, Staff, Datetime_Upload 
                            FROM TestSummary WHERE Datetime_Upload IS NULL
                            ''')
            query.exec_() 
            # 檢查並更新或新增到 MariaDB 資料表
            c=[]
            self.txtMsg_append("上傳測試總結記錄中...")
            urow=0
            while query.next():
                if self.StopTest: break
                urow+=1
                self.ledDataUpload.setText(f"上傳總結第{urow}筆")
                QApplication.processEvents()
                c.clear()
                for i in range(query.record().count()):
                    if i in [4,5,10, 11] and query.value(i)=='':
                        c.append('NULL')
                    else:
                        c.append(query.value(i))
                # 檢查 MariaDB 資料表 是否已經存在相同的記錄
                # query2 = QSqlQuery(self.MariaDB)
                cur1.execute(f"SELECT TestDocNo" +
                                f" FROM TestSummary " +
                                f" WHERE TestDocNo = '{c[0]}' ")            
                result1=cur1.fetchall()
                rcount=len(result1)
                # query3 = QSqlQuery(self.MariaDB)
                # self.MariaDB.transaction()
                UpDateTime=QDateTime.currentDateTime().toString(self.MyDTFormat)
                if rcount>0:
                    # 記錄已存在，使用更新方式   
                    strSQL=f"""UPDATE TestSummary SET TestDocNo = '{c[0]}', ScriptName = '{c[1]}', ScriptVersion = '{c[2]}', Workstation = '{c[3]}', 
                                CreateDatetime = '{c[4]}', LastTestDatetime = '{c[5]}', TestResult = '{c[6]}', 
                                DUTModel = '{c[7]}', DUTSerNo = '{c[8]}', DUTIDN = '{c[9]}', Temperature = {c[10]}, Humidity = {c[11]}, EquipList = '{c[12]}', 
                                Staff = '{c[13]}', Datetime_Upload = '{UpDateTime}'
                                WHERE TestDocNo = '{c[0]}'"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    cur2.execute(strSQL)                
                else:
                    # 記錄不存在，使用新增方式                
                    strSQL=f"""INSERT INTO TestSummary (TestDocNo, ScriptName, ScriptVersion, Workstation, CreateDatetime, LastTestDatetime, TestResult, 
                                DUTModel, DUTSerNo, DUTIDN, Temperature, Humidity, EquipList, Staff, Datetime_Upload) 
                                VALUES ('{c[0]}', '{c[1]}', '{c[2]}', '{c[3]}', '{c[4]}', '{c[5]}', '{c[6]}', '{c[7]}', '{c[8]}', '{c[9]}', {c[10]}, 
                                {c[11]}, '{c[12]}', '{c[13]}', '{UpDateTime}')"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    cur2.execute(strSQL)
                conn.commit()
                strSQLErr=query.lastError().text()
                if strSQLErr =="":
                    # query4 = QSqlQuery(self.db)
                    query4.prepare(f"UPDATE TestSummary SET Datetime_Upload = :c0 WHERE TestDocNo = '{c[0]}'")
                    query4.bindValue(':c0', UpDateTime)
                    query4.exec_()
                    self.db.commit()
                else:
                    self.txtMsg_append(f"SQL Error: \n{str(strSQLErr)}",Qt.red) 
                    # self.db.rollback()
            QApplication.processEvents()

            #===上傳TestLog資料表===
            query.prepare('''SELECT TestDocNo, ItemNo, SubItemNo, TestPointNo, Workstation, Staff, 
                            Measure_Main,  Measure_1, Measure_2, Judge, Comment,
                            Datetime_Start, Datetime_End, Datetime_Upload 
                            FROM TestLog WHERE Datetime_Upload IS NULL
                            ''')
            query.exec_() 
            # 檢查並更新或新增到 MariaDB 資料表
            c=[]
            self.txtMsg_append("上傳測試Log中...")
            urow=0
            while query.next():
                if self.StopTest: break
                urow+=1
                self.ledDataUpload.setText(f"上傳總結第{urow}筆")
                QApplication.processEvents()
                c.clear()
                for i in range(query.record().count()):
                    if i in [1,2,3,6,7,8,11,12,13] and query.value(i)=='':
                        c.append('NULL')
                    else:
                        c.append(query.value(i))
                # 檢查 MariaDB 資料表 是否已經存在相同的記錄
                sSQL=(f"SELECT TestDocNo, Datetime_Start" +
                                f" FROM TestLog " +
                                f" WHERE TestDocNo = '{c[0]}' and Datetime_Start= '{c[11]}'")
                sSQL=sSQL.replace("'NULL'","NULL")
                cur1.execute(sSQL)            
                result1=cur1.fetchall()
                rcount=len(result1)
                UpDateTime=QDateTime.currentDateTime().toString(self.MyDTFormat)
                if rcount>0:
                    # 記錄已存在，使用更新方式   
                    strSQL=f"""UPDATE TestLog SET TestDocNo = '{c[0]}', ItemNo = {c[1]}, SubItemNo = {c[2]}, TestPointNo={c[3]}, Workstation = '{c[4]}',
                                Staff='{c[5]}', Measure_Main={c[6]}, Measure_1={c[7]}, Measure_2={c[8]} , Judge='{c[9]}', Comment="{c[10]}",
                                Datetime_Start = '{c[11]}', Datetime_End = '{c[12]}', Datetime_Upload = '{UpDateTime}' 
                                WHERE TestDocNo = '{c[0]}' and Datetime_Start= '{c[11]}'"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    cur2.execute(strSQL)                
                else:
                    # 記錄不存在，使用新增方式                
                    strSQL=f"""INSERT INTO TestLog (
                            TestDocNo, ItemNo, SubItemNo, TestPointNo, Workstation, Staff, 
                            Measure_Main, Measure_1, Measure_2, Judge, Comment, Datetime_Start, Datetime_End, Datetime_Upload
                            ) 
                            VALUES ('{c[0]}', {c[1]}, {c[2]}, {c[3]}, '{c[4]}', '{c[5]}', {c[6]}, {c[7]}, {c[8]}, '{c[9]}', "{c[10]}", 
                            '{c[11]}', '{c[12]}' , '{UpDateTime}')"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    cur2.execute(strSQL)
                conn.commit()
                strSQLErr=query.lastError().text()
                if strSQLErr =="":
                    # query4 = QSqlQuery(self.db)
                    query4.prepare(f"UPDATE TestLog SET Datetime_Upload = :c0 WHERE TestDocNo = '{c[0]}' and Datetime_Start= '{c[11]}'")
                    query4.bindValue(':c0', UpDateTime)
                    query4.exec_()
                    self.db.commit()
                else:
                    self.txtMsg_append(f"SQL Error: \n{str(strSQLErr)}",Qt.red) 
                    # self.db.rollback()
            QApplication.processEvents()
            cur1.close()
            cur2.close()
            conn.close()
            self.txtMsg_append('上傳測試資料完成.')
            self.ledDataUpload.setText("")
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)

    #下載雲端測試腳本
    def DownloadTestScript(self): 
        try:            
            conn = pymysql.connect(
                user=self.MariaDB_User,
                password=self.MariaDB_Password,
                host=self.MariaDB_Host,
                port=self.MariaDB_port,
                database=self.MariaDB_Name)
            
            query = QSqlQuery(self.db)
            cur1=conn.cursor()

            #===下載雲端GPT測試腳本===  
            strSQL=f"""
                select Name, Version, ItemNo, SubItemNo, TestPointNo, ItemName, ItemName_2t, ItemName_3t,
                CondName_1, Cond_1, CondUnit_1,  CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3,
                Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit, 
                Cond_DUT, Cond_Equip, Cond_Fixture, Cond_Fixture_Other, Comment 
                from TestScript Where (Name, Version) in (
                select Name, max(Version) as Version from TestScript 
                where Name like 'GPT%'
                GROUP BY NAME
                ) 
                """            
            cur1.execute(strSQL)
            result1=cur1.fetchall()
            if len(result1)>0:
                urow=0
                self.txtMsg_append(f"刪除本地腳本...")
                query.prepare('''
                    delete from TestScript
                    ''')
                query.exec_()       #刪除本地端腳本
                self.db.commit()
                self.txtMsg_append(f"下載雲端腳本中...")
                QApplication.processEvents()
                c=[]    
                for r in result1:
                    if self.StopTest: break
                    urow+=1
                    self.ledDataUpload.setText(f"下載雲端腳本第{urow}筆")
                    QApplication.processEvents()
                    c.clear()
                    for i in range(len(r)):
                        if r[i]=='':
                            c.append('NULL')
                        else:
                            c.append(r[i])
                            
                    strSQL=f"""INSERT INTO TestScript (Name, Version, ItemNo, SubItemNo, TestPointNo, ItemName, ItemName_2t, ItemName_3t,
                                CondName_1, Cond_1, CondUnit_1,  CondName_2, Cond_2, CondUnit_2, CondName_3, Cond_3, CondUnit_3,
                                Expect, ExpectUnit, LowLimit, UpLimit, Deviation_Expect, Deviation_LowLimit, Deviation_UpLimit, 
                                Cond_DUT, Cond_Equip, Cond_Fixture, Cond_Fixture_Other, Comment) 
                                VALUES ('{c[0]}', '{c[1]}', {c[2]}, {c[3]}, {c[4]}, '{c[5]}', '{c[6]}', '{c[7]}', '{c[8]}', '{c[9]}', '{c[10]}', 
                                '{c[11]}', '{c[12]}', '{c[13]}', '{c[14]}', '{c[15]}', '{c[16]}', {c[17]}, '{c[18]}', {c[19]}, {c[20]}, 
                                {c[21]}, {c[22]}, {c[23]}, '{c[24]}', '{c[25]}', '{c[26]}', '{c[27]}', '{c[28]}')"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    strSQL=strSQL.replace("\n","")                       
                    query.prepare(strSQL)
                    query.exec_()
                    self.db.commit()
                    if query.lastError().text() !='':
                        print(f"{inspect.currentframe().f_code.co_name} SQL錯誤訊息:\n{query.lastError().text()}")  
            
            #===下載雲端GPT測試腳本與機種對照表===  
            strSQL=f"""
                select Model, Script 
                from ScriptRef 
                where Script like 'GPT%' 
                """            
            cur1.execute(strSQL)
            result1=cur1.fetchall()
            if len(result1)>0:
                urow=0
                self.txtMsg_append(f"刪除本地機種腳本對照表...")
                query.prepare('''
                    delete from ScriptRef
                    ''')
                query.exec_()       #刪除本地端機種腳本對照表
                self.db.commit()
                self.txtMsg_append(f"下載雲機種腳本對照表中...")
                QApplication.processEvents()
                c=[]    
                for r in result1:
                    if self.StopTest: break
                    urow+=1
                    self.ledDataUpload.setText(f"下載雲機種腳本對照表第{urow}筆")
                    QApplication.processEvents()
                    c.clear()
                    for i in range(len(r)):
                        if r[i]=='':
                            c.append('NULL')
                        else:
                            c.append(r[i])
                            
                    strSQL=f"""INSERT INTO ScriptRef (Model, Script) 
                                VALUES ('{c[0]}', '{c[1]}')"""
                    strSQL=strSQL.replace("'NULL'","NULL")
                    strSQL=strSQL.replace("\n","")                       
                    query.prepare(strSQL)
                    query.exec_()
                    self.db.commit()
                    if query.lastError().text() !='':
                        print(f"{inspect.currentframe().f_code.co_name} SQL錯誤訊息:\n{query.lastError().text()}")                                  
             
            QApplication.processEvents()
            # cur1.close()            
            # conn.close()
            self.txtMsg_append('下載雲端測試腳本及機種腳本對照表完成.')
            self.ledDataUpload.setText("")

            #===新增治具路徑控制表===  
            strSQL=f"""
                select PName, PDesc, Port0, Port1, Port2, Port3, Port4, Port5 
                from FixturePath 
                """            
            cur1.execute(strSQL)
            result1=cur1.fetchall()
            if len(result1)>0:
                urow=0
                # self.txtMsg_append(f"刪除本地機種腳本對照表...")
                # query.prepare('''
                #     delete from ScriptRef
                #     ''')
                # query.exec_()       #刪除本地端機種腳本對照表
                # self.db.commit()
                self.txtMsg_append(f"新增治具路徑控制表項目...")
                QApplication.processEvents()
                c=[]    
                for r in result1:
                    if self.StopTest: break
                    
                    c.clear()
                    for i in range(len(r)):
                        if r[i]=='':
                            c.append('NULL')
                        else:
                            c.append(r[i])
                    strSQL=f""" select PName from FixturePath where PName='{c[0]}'
                            """        
                    query.prepare(strSQL)
                    query.exec_()
                    rnum=0
                    while(query.next()):
                        rnum+=1
                        break
                    if rnum==0:
                        urow+=1
                        self.ledDataUpload.setText(f"新增治具路徑表第{urow}筆")
                        QApplication.processEvents()
                        strSQL=f"""INSERT INTO FixturePath (PName, PDesc, Port0, Port1, Port2, Port3, Port4, Port5) 
                                    VALUES ('{c[0]}', '{c[1]}', '{c[2]}', '{c[3]}', '{c[4]}', '{c[5]}', '{c[6]}', '{c[7]}')"""
                        strSQL=strSQL.replace("'NULL'","NULL")
                        strSQL=strSQL.replace("\n","")                       
                        query.prepare(strSQL)
                        query.exec_()
                        self.db.commit()
                        if query.lastError().text() !='':
                            print(f"{inspect.currentframe().f_code.co_name} SQL錯誤訊息:\n{query.lastError().text()}")                                  
            #===新增治具路徑校正值項目===  
            strSQL=f"""
                select PName, Cond_1, Cond_1_unit, Cond_2, Cond_2_unit, CalValue, CalValue_unit 
                from FixtureCal
                """            
            cur1.execute(strSQL)
            result1=cur1.fetchall()
            if len(result1)>0:
                urow=0
                # self.txtMsg_append(f"刪除本地機種腳本對照表...")
                # query.prepare('''
                #     delete from ScriptRef
                #     ''')
                # query.exec_()       #刪除本地端機種腳本對照表
                # self.db.commit()
                self.txtMsg_append(f"新增治具路徑校正值項目...")
                QApplication.processEvents()
                c=[]    
                for r in result1:
                    if self.StopTest: break
                    
                    c.clear()
                    for i in range(len(r)):
                        if r[i]=='':
                            c.append('NULL')
                        else:
                            c.append(r[i])
                    strSQL=f""" select PName from FixtureCal where PName='{c[0]}'
                            """        
                    query.prepare(strSQL)
                    query.exec_()
                    rnum=0
                    while(query.next()):
                        rnum+=1
                        break
                    if rnum==0:
                        urow+=1
                        self.ledDataUpload.setText(f"新增治具路徑校正值第{urow}筆")
                        QApplication.processEvents()
                        strSQL=f"""INSERT INTO FixtureCal (PName, Cond_1, Cond_1_unit, Cond_2, Cond_2_unit, CalValue, CalValue_unit) 
                                    VALUES ('{c[0]}', '{c[1]}', '{c[2]}', '{c[3]}', '{c[4]}', '{c[5]}', '{c[6]}')"""
                        strSQL=strSQL.replace("'NULL'","NULL")
                        strSQL=strSQL.replace("\n","")                       
                        query.prepare(strSQL)
                        query.exec_()
                        self.db.commit()
                        if query.lastError().text() !='':
                            print(f"{inspect.currentframe().f_code.co_name} SQL錯誤訊息:\n{query.lastError().text()}")                                 
             
            QApplication.processEvents()
            cur1.close()            
            conn.close()
            self.txtMsg_append('下載雲端測試腳本及機種腳本對照表完成.')
            self.ledDataUpload.setText("")
            self.delay(1)


        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
            

    #顥示設備列表
    def showEquip(self):    
        try:
            self.db =  QSqlDatabase.addDatabase('QSQLITE')  #建立資料庫
            self.tbView.setWindowTitle("設備列表")
            if self.db.databaseName()!=self.DBPth:
                self.db.setDatabaseName(self.DBPth)         #設定資料庫路徑名稱           
            if not self.db.open():  self.db.open()          #資料庫没有開啟, 則開啟之            
            if not self.db.open():                          #資料庫還是没有開啟, 返叵      
                return False
            self.model=CustomSqlModel()        
            self.model.setTable('TestEquip')        
            self.tbView.show()
            self.model.select()
            if self.model.rowCount()!=0:
                self.model.setHeaderData(1, Qt.Horizontal, "名稱")
                self.model.setHeaderData(2, Qt.Horizontal, "型號")
                self.model.setHeaderData(3, Qt.Horizontal, "連接口")
                self.model.setHeaderData(4, Qt.Horizontal, "鮑率")
                self.model.setHeaderData(5, Qt.Horizontal, "查詢關鍵字")
                self.model.setHeaderData(6, Qt.Horizontal, "是否連接")                
                self.tbView.setModel(self.model)
                self.tbView.hideColumn(0)
                #self.tbView.setColumnWidth(0,40)
                self.tbView.setColumnWidth(1,120)
                self.tbView.setColumnWidth(2,120)
                self.tbView.setColumnWidth(3,150)
                font=self.tbView.font()
                font.setPointSize(11)
                # font.setPixelSize(11)
                self.tbView.setFont(font) 
            else:
                self.tbView.setModel(self.model)
                print("找不到設備列表, 請確認資料庫...")
        except Exception as e:            
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        
    #設定設備連線相關資訊---------------------------------------------------------------------------------------------------------------------------
    def setDevInfo(self):
        QtWidgets.QApplication.processEvents() 
        try:       
            for i in range(0,self.model.rowCount()):            
                rdev=self.model.record(i) 
                if rdev.value('BaudRate') not in self.BRList:
                    self.BRList.append(rdev.value('BaudRate'))             
                if rdev.value('Name')=='待測機_RS232':   
                    dev=DUT_RS232                
                elif rdev.value('Name')=='待測機_USB':
                    #dev=DUT_USB  
                    dev=self.DUT
                elif rdev.value('Name')=='待測機_GPIB':
                    dev=DUT_GPIB                
                elif rdev.value('Name')=='高壓表':
                    dev=HVM                
                elif rdev.value('Name')=='DMM':
                    dev=DMM                
                elif rdev.value('Name')=='電流表':
                    dev=PCS                
                elif rdev.value('Name')=='DIO控制器':
                    dev=DIO
                elif rdev.value('Name')=='直流電源':
                    dev=PSW
                
                dev.VisaDescription=rdev.value('PortDescr')
                dev.Baudrate=rdev.value('BaudRate')
                dev.IDN_Keyword=rdev.value('KeyWord')
                dev.instrConnected=rdev.value('Connected')
                
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
    #----------------------------------------------------------------------------------------------------------------------------------------------------

    #擷取資籵, 顯示測試選項
    def showPathItem(self):
        try:
            self.cb=[]              #測試項目核選方塊
            self.cb.clear() 
            self.tj=[]
            self.tj.clear()
            r=self.GLayout.rowCount()
            if r>0:     #清除Grid Layout的控制項
                for i in reversed(range(self.GLayout.count())): 
                    self.GLayout.itemAt(i).widget().deleteLater()
            QtWidgets.QApplication.processEvents()
            self.rsItem=self.getTestItems() #取測試項目名稱
            if len(self.rsItem)>0:
                self.cboScript.setCurrentText(self.rsItem[0][2])
            for i in range(0,len(self.rsItem)):            
                self.cb.append(QtWidgets.QCheckBox(self))
                self.cb[i].setObjectName("chkItem_"+str(i))
                self.cb[i].setText(str(self.rsItem[i][0]) + "."+str(self.rsItem[i][1]))
                tn=str(self.cb[i].text()).split('.')
                font = self.cb[i].font()
                self.cb[i].setChecked(False)
                self.cb[i].setEnabled(False)
                #self.cb[i].setStyleSheet('color:gray')
                for k in self.OKItem:            #若為程式已經完成的項目, 預設值為勾選
                    if fnmatch(tn[1], k):                    
                        self.cb[i].setEnabled(True)
                        self.cb[i].setChecked(True)
                        #self.cb[i].setStyleSheet('color:black')                    
                        break            
                font.setPointSize(11)
                self.cb[i].setFont(font)
                self.tj.append(QLabel(self))
                self.tj[i].setObjectName("lblResult_"+str(i))
                QtWidgets.QApplication.processEvents()
                self.tj[i].setFont(font)
                self.tj[i].setStyleSheet('color:black')

                self.GLayout.addWidget(self.cb[i],i,0)
                self.GLayout.addWidget(self.tj[i],i,1)
                QtWidgets.QApplication.processEvents()
                
            self.gboxTestItems.setLayout(self.GLayout)
            self.scrollArea.setWidget(self.gboxTestItems)
            self.scrollArea.setWidgetResizable(True) 
            QtWidgets.QApplication.processEvents()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)

    #擷取資料, 顯示治具選項
    def showPathFix(self): 
        # print(self.getTestItems()[0]['ItemNo'])
        try:
            self.cf=[]              #治具項目核選方塊
            self.cf.clear()
            r=self.GLayout_F.rowCount()
            if r>0:     #清除Grid Layout的控制項
                for i in reversed(range(self.GLayout_F.count())): 
                    self.GLayout_F.itemAt(i).widget().deleteLater()
            QtWidgets.QApplication.processEvents()
            self.rsFix=self.getFixCond() #取測試治具條件   
            for i in range(0,len(self.rsFix)):            
                self.cf.append(QtWidgets.QCheckBox(self))
                self.cf[i].setObjectName("chkFix_"+str(i))
                self.cf[i].setText(str(self.rsFix[i][0]))            
                self.cf[i].setChecked(True)
                font = self.cf[i].font()
                font.setPointSize(11)
                self.cf[i].setFont(font)
                QtWidgets.QApplication.processEvents()   
                self.GLayout_F.addWidget(self.cf[i],i,0)
            self.gboxFix.setLayout(self.GLayout_F)
            self.scrollArea_3.setWidget(self.gboxFix)
            self.scrollArea_3.setWidgetResizable(True)
            QtWidgets.QApplication.processEvents()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
    #按下匯出測試記錄按鈕
    def on_btnExportRec_Click(self):
        window.TRecord.table_export()
    #按下設備確認
    def on_btnConDev_Click(self): 
        self.onlyStopEnable()   #設為只有STOP可按           
        self.StopTest=False
        self.ConDevCheck()
        # self.setDIOPinMode_out() #設定MEGA 2560 Pin mode為OUTPUT
        self.onlyStopDisable()  #設為只有STOP不可按

    def ConDevCheck(self):
        t1=time() 
        #self.DBPth=self.ledDBPath.text()        #取得新資料庫路徑
        #window.DBPth=self.DBPth
        self.txtMsg.setText('設備連線確認中...')
        QtWidgets.QApplication.processEvents()
        try:
            dev=MyInstr.Instrument()
            dev.instrDelayTime_s=0.5
            dev.instrTimeouts_ms=500
            tupDevices=dev.rm.list_resources("?*")
            print(str(tupDevices))
            print(dev.VisaDescription)
            PList=[]
            FBRList=[]
            FAgain=False
            self.EquipsBeUsed=''
            self.showEquip()
            if self.model.rowCount()==0:
                self.txtMsg_append('找不到設備列表, 請確認資料庫...',Qt.red) 
            i=0
            for d in tupDevices:
                PList.append(d)
            self.SkipPort.clear()
            self.SkipPort=self.txtSkipPort.toPlainText().split('\n')
            for i in range(0,len(self.SkipPort)):
                self.SkipPort[i]=str(self.SkipPort[i].strip().upper()) #將SKIPPORT名稱轉大寫
            
            #資料表的資料, 確認記載的Port是否有連接
            for i in range(0,self.model.rowCount()):  
                if self.StopTest: break
                rdev=self.model.record(i) 
                dev.VisaDescription=rdev.value('PortDescr')
                dev.Baudrate=rdev.value('BaudRate')
                dev.IDN_Keyword=rdev.value('KeyWord') 
                if dev.VisaDescription not in tupDevices:
                    self.txtMsg_append(f"{rdev.value('Name')}, {dev.VisaDescription}, 不在資源列表中",Qt.red)
                    rdev.setValue('Connected',0)
                    if dev.Baudrate not in FBRList:
                        FBRList.append(dev.Baudrate)
                    FAgain=True     #標装有機器没有找到
                    self.model.setRecord(i,rdev)
                    self.model.submitAll()
                    continue

                if fnmatch(rdev.value('KeyWord'),'*MEGA*'):
                    dev.instrDelayTime_s=0.5          #MEGA 2560 Port open後的反應時間會較慢一些, 故延長 0.25*5=1.25秒
                else:
                    dev.instrDelayTime_s=0.1 
                print(dev.VisaDescription)        
                if dev.checkPort():
                    print(dev.VisaDescription)
                    
                    dev.openPort()
                    id=dev.get_IDN().replace('\n','') 
                    self.EquipsBeUsed+='{"Name":"' + rdev.value('Name')+ '","IDN":"' + id +'"}'
                    self.txtMsg_append(f"找到{rdev.value('Name')}=>Port: {dev.VisaDescription} IDN: {id}")
                    rdev.setValue('Connected',1)  
                    rdev.setValue('PortDescr',dev.VisaDescription)
                    dev.closePort()
                    if dev.VisaDescription in PList: PList.remove(dev.VisaDescription)                
                else:
                    self.txtMsg_append(f"{rdev.value('Name')}在{dev.VisaDescription}没找到",Qt.red)
                    rdev.setValue('Connected',0)
                    FAgain=True     #標装有機器没有找到
                    if dev.Baudrate not in FBRList:
                        FBRList.append(dev.Baudrate)
                self.model.setRecord(i,rdev)
                self.model.submitAll()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        self.showEquip()
        QtWidgets.QApplication.processEvents()
        try:
            #還没確認的Port再掃一次
            for j in PList:            
                if FAgain==False: break
                if self.StopTest: break
                skipJ=False
                # if str(j).strip().upper() in self.SkipPort: continue    #忽略檢查Skip port
                for k in range(0,len(self.SkipPort)):
                    if fnmatch(j,self.SkipPort[k]): 
                        skipJ=True
                        break
                if skipJ: continue    #忽略檢查Skip port
                self.txtMsg_append(f"再次確認{j}")            
                dev.VisaDescription=j
                dev.instrTimeouts_ms=500
                dev.instrDelayTime_s=0.5    #因有些設備反應較慢, 如MEGA 2560全掃時放慢
                q=''              
                if str(dev.VisaDescription).lower().find('asrl')>=0:
                    for k in FBRList:
                        if self.StopTest: break
                        if k=='':
                            continue
                        dev.Baudrate=k 
                        dev.openPort()
                        q=dev.get_IDN()                        
                        self.txtMsg_append(f"以Baud rate: {k} 確認, 讀到IDN: {q}")
                        dev.closePort()                    
                        if q !='': break
                else:
                    dev.instrDelayTime_s=0.1
                    dev.openPort()
                    q=dev.get_IDN()                
                    self.txtMsg_append(f"讀到IDN: {q}")
                    dev.closePort()
                QtWidgets.QApplication.processEvents()
                    
                for i in range(0,self.model.rowCount()): 
                    if self.StopTest: break           
                    rdev=self.model.record(i)
                    if q.lower().find(str(rdev.value('Keyword')).lower())>=0:
                        self.EquipsBeUsed+='{"Name":"' + rdev.value('Name')+ '","IDN":"' + q +'"}'
                        self.txtMsg_append(f"找到{rdev.value('Name')}, Port:{dev.VisaDescription}, Baud rate:{dev.Baudrate}",Qt.blue)
                        rdev.setValue('Connected',1)  
                        rdev.setValue('PortDescr',dev.VisaDescription)
                        rdev.setValue('BaudRate',dev.Baudrate)
                        self.model.setRecord(i,rdev)
                        self.model.submitAll()
                        QtWidgets.QApplication.processEvents()
                        break
                    QtWidgets.QApplication.processEvents()            
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        self.showEquip()        
        self.setDevInfo()
        print(DUT_USB.VisaDescription)
        print(DUT_USB.VisaDescription)
        if self.StopTest:
            self.txtMsg_append('[停止]被按, 停止後續設備連線確認...',Qt.red)            
        elif self.model.rowCount()>0:
            self.txtMsg_append(f"設備連線確認完成, 花費{str(round(time()-t1,1))}秒",Qt.blue)
        else:
            self.txtMsg_append('找不到設備列表, 請確認資料庫...',Qt.red)         
            
    #附加文字到txtMsg
    def txtMsg_append(self,s="",TextColor=Qt.black):        
        self.txtMsg.setTextColor(TextColor)
        self.txtMsg.append(s)
        self.txtMsg.verticalScrollBar().setValue(self.txtMsg.verticalScrollBar().maximum()) #垂宜捲到最下
        QtWidgets.QApplication.processEvents()

    def onExitClick(self):      
        self.timer1.stop()
        self.StopTest=True
        self.delay(0.2)
        self.close()
        MainClose()
    
    def delay(self,delaySec=0.0,start_monotonic=''):
        try:
            wt=0
            if start_monotonic=='':
                st=monotonic()
            else:
                st=start_monotonic            
            while(wt<delaySec):
                wt=monotonic()-st
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈
                sleep(0.001)
        except Exception as e:
            print(str(e))
        

    #設定除了STOP外的按鍵都可按
    def onlyStopDisable(self):
        subProcName="onlyStopDisable"
        try:            
            btns=self.findChildren(QPushButton)
            for bt in btns:
                font = bt.font()            
                font.setPointSize(11)
                bt.setFont(font)
                bt.setDisabled(False)
                bt.setStyleSheet('color:black')
                

            self.btnStop.setDisabled(True)
            self.btnStop.setStyleSheet('color:gray')
            self.btnStart.setStyleSheet('color:blue')
            self.chkFixManual.setEnabled(True)
            self.chkFixOptimize.setEnabled(True)
            self.chkHVMManual.setEnabled(True)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
    
    
    #設定只有Stop按鍵可以按
    def onlyStopEnable(self): 
        subProcName="onlyStopEnable"
        try:
            btns=self.findChildren(QPushButton)
            for bt in btns:
                font = bt.font()            
                font.setPointSize(11) 
                bt.setFont(font) 
                bt.setDisabled(True)
                bt.setStyleSheet('color:gray')
                    
            self.btnStop.setDisabled(False)
            self.btnStop.setStyleSheet('color:red')        
            self.chkFixManual.setEnabled(False)
            self.chkFixOptimize.setEnabled(False)
            self.chkHVMManual.setEnabled(False)
            self.btnStop.setFocus()
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)

    
    #按下開始鍵
    def onStartClick(self):
        TH=recordHeader()
        subProcName="onStartClick"   
        self.ledDataUpload.setText("")
        self.SwitchPath=''
        try:     
            self.timer1.start(1000)
            self.stimer=monotonic()
            #self.DBPth=self.ledDBPath.text()        #取得新資料庫路徑
            datetime = QDateTime.currentDateTime()
            self.ltxtStartTime.setText(datetime.toString(self.MyDTFormat))
            self.txtMsg.setText('')
            self.onlyStopEnable() #設為只有STOP可按  
            self.getSpecTestRow()   #取得指定要測試的ROW  
            if self.cboReportNo.currentText().find(self.ltxtDUTSn.text())<0:
                self.onNewReport()
            self.StopTest=False
            self.clearTestJudge()
            TDocNo=self.cboReportNo.currentText()   #測試文件編號
            #------------------------------------------------------------------------------------------------------------
            if fnmatch(self.cboScript.currentText(),'PEL-3*'): #依腳本, 設定DUT物件
                self.DUT=MyInstr.PEL()
            elif fnmatch(self.cboScript.currentText(),'*120*') or fnmatch(self.cboScript.currentText(),'*150*'): #依腳本, 設定DUT物件
                self.DUT=MyInstr.GPT10000()
            else:
                self.DUT=MyInstr.GPT()
            #-------------------------------------------------------------------------------------------------------------
         
            
            #若有指定測試列,把指定測試列使用的治具路徑記錄在self.SpecTRowFix
            if self.SpecTRow!=[]:                
                self.SpecTRowFix.clear()
                for i in range(0,window.TRecord.model.rowCount()):             
                    QtWidgets.QApplication.processEvents()
                    if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈
                    if i in self.SpecTRow:
                        r=window.TRecord.model.record(i)
                        if r.value(TH.col_Cond_Fixture) not in self.SpecTRowFix:
                            self.SpecTRowFix.append(r.value(TH.col_Cond_Fixture))
            
            self.ConDevCheck()   #設備進行連線
                    
            # try:
            #     DIO.openPort()
            # except: pass
            # self.setDIOPinMode_out() #設定MEGA 2560 Pin mode為OUTPUT
            self.FirstExecPQC=True      #設定將為按下按鈕後, 第一次執行ExecPQC
            self.checkedFix.clear()
            for i in self.cf:           #記錄有勾選的治具
                if i.isChecked():
                    self.checkedFix.append(i.text())
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        
        try:
            if self.chkFixOptimize.isChecked():     
                #有勾選治具控制優化-依治具路徑順序測試 (減少換線次數)
                for j in range(0,len(self.cf)):
                    if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈
                    if not(self.cf[j].isChecked()): #若此項治具没有勾選, 跳到下個治具
                        QtWidgets.QApplication.processEvents()
                        continue
                    if not(self.isFixUsed(self.cf[j].text())):continue  #若勾選的測項殳有用到此治具, 跳到下個治具
                    
                    if self.chkFixManual.isChecked():   #若有勾選手動換線
                        self.SwitchPath=self.cf[j].text()
                        if self.SpecTRow!=[]:
                            if self.SwitchPath not in self.SpecTRowFix: continue #若有指定測試列, 但治具路徑不在指定測試列內, 略過此次換線的測試

                        reply = QMessageBox.information(self, '手動換線', f"請手動換線至\n{self.SwitchPath}", 
                                QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
                        if reply == QMessageBox.Cancel:
                            continue    #没有換線成功, 略過此次換線的測試
                        
                    for i in range(0,len(self.cb)): #從頭開始跑測項
                        if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈
                        # print(self.cb[i].text())
                        rect_cb=self.cb[i].geometry()
                        self.scrollArea_2.ensureVisible(0,rect_cb.y())
                        font = self.tj[i].font()
                        font.setBold(True)
                        self.tj[i].setFont(font)
                        if self.cb[i].isChecked():   #若有勾選此測項, 執行測項測試  
                            ItemName=self.cb[i].text()
                            IName=ItemName[ItemName.find('.')+1:]
                            if not(self.isFixUsed(self.cf[j].text(),IName)):continue #若此測項没有用到此治具, 跳到下個測項 

                            self.tj[i].setText('測試中')                
                            self.tj[i].setStyleSheet('color:green')                        
                            QtWidgets.QApplication.processEvents()                
                            sJudge=self.ExecPQC(self.cb[i].text(),self.cf[j].text())   #增加判斷治具條件執行測試
                            self.delay(0.1)
                            if str(sJudge).lower() in ['pass','ok']:
                                self.tj[i].setStyleSheet('color:blue')
                                self.tj[i].setText(sJudge)
                            elif str(sJudge).lower() in ['fail','ng']:
                                self.tj[i].setStyleSheet('color:red')
                                self.tj[i].setText(sJudge)
                            else:
                                self.tj[i].setText("")               
                            QtWidgets.QApplication.processEvents()
            else:        
                #依一般測項順序測試
                for i in range(0,len(self.cb)):
                    if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈
                    rect_cb=self.cb[i].geometry()
                    self.scrollArea_2.ensureVisible(0,rect_cb.y())
                    font = self.tj[i].font()
                    font.setBold(True)
                    self.tj[i].setFont(font)
                    if self.cb[i].isChecked():                
                        self.tj[i].setText('測試中')                
                        self.tj[i].setStyleSheet('color:green')
                        QtWidgets.QApplication.processEvents()
                        #print(self.DUT.VisaDescription)
                        #print(self.VisaDescription)                
                        sJudge=self.ExecPQC(self.cb[i].text())
                        self.delay(0.1)
                        if str(sJudge).lower() in ['pass','ok']:
                            self.tj[i].setStyleSheet('color:blue')
                            self.tj[i].setText(sJudge)
                        elif str(sJudge).lower() in ['fail','ng']:
                            self.tj[i].setStyleSheet('color:red')
                            self.tj[i].setText(sJudge)
                        else:
                            self.tj[i].setText("")                
                        QtWidgets.QApplication.processEvents()
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)            
        self.timer1.stop()
        self.onTimerTick()
        try:
            DIO.closePort()
            DMM.closePort()
            self.DUT.closePort()
            HVM.closePort()
            PCS.closePort()
        except: pass
        if self.StopTest: 
            self.txtMsg_append('[停止]被按, 停止後續測試...',Qt.red)
        else:
            self.txtMsg_append('測試程序完了.',Qt.blue)
        self.updateTestSummary()
        self.onlyStopDisable() #測試結束, 其他都可按, 只有STOP不能
        self.UploadTestData()           #上傳測試資料
        self.ltxtDUTModel.setFocus()
        self.ltxtDUTModel.selectAll()
        
    
    
    #確認引數的治具條件是否在勾選項目(或指定的項目)有使用到            
    def isFixUsed(self, FixUse, ItemName=''):
        TH=recordHeader()
        TI=[]
        TI.clear()
        u=False
        if ItemName=='':
            for i in range(0,len(self.cb)):
                if self.cb[i].isChecked():
                    cbs=self.cb[i].text().split('.')
                    TI.append(cbs[1])
        else:
            TI.append(ItemName)
        # print(TI)
        for i in range(0,window.TRecord.model.rowCount()):            
            r=window.TRecord.model.record(i)
            # print(f"r.value(TH.col_ItemName)=>{r.value(TH.col_ItemName)}")
            if r.value(TH.col_ItemName) not in TI: 
                continue #不是此測項, 不執行此圈
            if r.value(TH.col_Cond_Fixture)==FixUse:
                u=True
                break        
        return u

    def clearTestJudge(self):
        for i in range(0,len(self.tj)):            
            #print(i.objectName())
            #self.tj[i].setText('未測')
            self.tj[i].setText('')
            self.tj[i].setStyleSheet('color:black')
            QtWidgets.QApplication.processEvents()
    
    #取得指定的測試列
    def getSpecTestRow(self):
        listA=[]
        listB=[]        
        self.SpecTRow.clear()
        listA.clear()        
        s1=self.txtSpecTRow.toPlainText().strip().replace('\n',',')
        if s1!='': 
            listA=s1.split(',')
            for i in listA:
                i=str(i).strip()
                if i=='': continue
                if i.find('-')>0:                    
                    listB.clear()
                    listB=i.split('-')
                    if len(listB) >=2:
                        listB[0]=listB[0].strip()
                        listB[1]=listB[1].strip()
                        try:
                            j1=int(listB[0])
                            j2=int(listB[1])
                            for j in range(j1,j2+1):
                                self.SpecTRow.append(j-1)
                        except Exception as e:
                            self.txtMsg_append(str(e),Qt.red)
                else:
                    try:
                        j=int(i)
                        self.SpecTRow.append(j-1)
                    except Exception as e:
                            self.txtMsg_append(str(e),Qt.red)

    def onStopClick(self):
        self.timer1.stop()
        self.StopTest=True
        # self.txtMsg.setText(self.txtMsg.toPlainText()+'\n[停止]被按下了\n')        
    
    #顯示已花費時間
    def onTimerTick(self):        
        spt=monotonic()-self.stimer
        m=int(spt/60)
        # ss=round(spt%60)
        ss=float(round(spt-m*60,1))
        self.ltxtSpendTime.setText(f'{m} 分 {ss} 秒')

    #測試腳本變更
    def oncboScriptChanged(self):
        print('oncboScriptChanged exec.')
        #self.showPathItem()
    
    def oncboReportNoChanged(self):
        print('oncboReportNoChanged exec.')        
        self.showPathItem()
        self.showPathFix()
        # print(f'self.cboReportNo.currentText()=>{self.cboReportNo.currentText()}')
        showTestResult()    #顯示測試表格及結果
        for i in range(0,len(self.cb)):
            s=self.cb[i].text()
            s=s[s.find('.')+1:]
            sJudge=self.JudgeResult(s)
            if str(sJudge).lower() in ['pass','ok']:
                self.tj[i].setStyleSheet('color:blue')
                self.tj[i].setText(sJudge)
            elif str(sJudge).lower() in ['fail','ng']:
                self.tj[i].setStyleSheet('color:red')
                self.tj[i].setText(sJudge) 
            else:
                self.tj[i].setText('')               
            QtWidgets.QApplication.processEvents()

    def on_ledDBPath_textChanged(self):
        try:
            s=self.ledDBPath.text().strip()
            window.setDBPth(s)
            self.DBPth=s
        except Exception as e:
            self.txtMsg_append(f"on_ledDBPath_textChanged Error:\n{str(e)}",Qt.red)

    def on_txtSpecTRow_textChanged(self):
        try:
            font = self.txtSpecTRow.font()
            font.setBold(True)
            self.txtSpecTRow.setFont(font)
            self.txtSpecTRow.setTextColor(color_Magenta)
        except Exception as e:
            self.txtMsg_append(f"on_txtSpecTRow_textChanged Error:\n{str(e)}",Qt.red)

    def closeEvent(self, event):
        self.StopTest=True
        self.delay(0.05)
        self.saveParameter()
        self.delay(0.05)
        event.accept()

    def saveParameter(self):
        self.settings.setValue("txtSkipPort", str(self.txtSkipPort.toPlainText()).strip())
        self.settings.setValue("DBPath", self.ledDBPath.text().strip())
        self.settings.setValue("Inspector", self.ltxtInspector.text().strip())
        self.settings.setValue("ltxtDUTModel", self.ltxtDUTModel.text().strip())
        self.settings.setValue("ltxtDUTSn", self.ltxtDUTSn.text().strip())
        if self.chkFixOptimize.isChecked():
            self.settings.setValue("chkFixOptimize", 1)
        else:
            self.settings.setValue("chkFixOptimize", 0)
        if self.chkFixManual.isChecked():
            self.settings.setValue("chkFixManual", 1)  
        else:
            self.settings.setValue("chkFixManual", 0)
        if self.chkHVMManual.isChecked():
            self.settings.setValue("chkHVMManual", 1)  
        else:
            self.settings.setValue("chkHVMManual", 0)

    def keyPressEvent(self, e):
        if e.key()  in [Qt.Key_Return, Qt.Key_Enter ]:  #Qt.Key_Return=>英文區的Enter鍵, Qt.Key_Enter=>數字區的Enter鍵
            if self.ltxtDUTSn.hasFocus():
                if len(self.ltxtDUTSn.text())>=6 and len(self.ltxtDUTSn.text())<=10:
                    self.getSNDocNo()
                    self.btnStart.setFocus()
                else:
                    QMessageBox.warning(self,"序號長度不符", "序號長度預設為6-10碼,\n請確認序號後, 重新輸入.")
                    self.ltxtDUTSn.setFocus()
                    self.ltxtDUTSn.selectAll()
                QtWidgets.QApplication.processEvents()
            elif self.ltxtDUTModel.hasFocus():
                if len(self.ltxtDUTModel.text())>=6 and len(self.ltxtDUTModel.text())<=15:
                    if self.selectScript():
                        self.ltxtDUTSn.setFocus()
                        self.ltxtDUTSn.selectAll()
                    else:
                        self.ltxtDUTModel.setFocus()
                        self.ltxtDUTModel.selectAll()
                else:
                    QMessageBox.warning(self,"型號長度不符", "型號長度預設為6-15碼,\n請確認型號後, 重新輸入.")
                    self.ltxtDUTModel.setFocus()
                    self.ltxtDUTModel.selectAll()
                QtWidgets.QApplication.processEvents()
            elif self.ltxtInspector.hasFocus():
                if self.ltxtInspector.text()!='admin':
                    # self.btnProTest.setVisible(False)
                    self.btnProTest.setVisible(True)                    
                else:
                    self.btnProTest.setVisible(True)
                self.ltxtDUTModel.setFocus()
                self.ltxtDUTModel.selectAll()
            elif self.btnStart.hasFocus():
                self.btnStop.setFocus()
                self.onStartClick()                
            elif self.btnStop.hasFocus():
                self.onStopClick()
            elif self.btnNewReport.hasFocus():
                self.onNewReport()
            elif self.btnExit.hasFocus():
                self.onExitClick()
            elif self.btnConDev.hasFocus():
                self.on_btnConDev_Click()
            elif self.btnMModify.hasFocus():
                self.on_btnMModify_Click()
            elif self.ledDBPath.hasFocus():
                self.checkDBField()
                self.getTestScript()
                self.selectScript()
                self.showPathItem()
                self.showPathFix()
                self.showEquip()            
                self.setDevInfo()
            else:
                print('其他欄輸入 Return或Enter')
        elif e.key() in [Qt.Key_Tab]:
            if self.txtSpecTRow.hasFocus():
                self.btnMModify.setFocus()


    #選取測試腳本
    def selectScript(self):
        subProcName="selectScript"
        found=False
        s=self.ltxtDUTModel.text()
        try:
            AteDB.ConnectSQLite(self.DBPth)            
            sSql=f"select distinct Script from ScriptRef where Model='{s}'"
            rsScript=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            if len(rsScript)>0: 
                m=str(rsScript[0][0])
            else:
                m=''         

            for i in range(0,self.cboScript.count()):
                if m==self.cboScript.itemText(i):
                    self.cboScript.setCurrentIndex(i)
                    self.txtMsg.setText('')
                    found=True
                    break
            if not found:
                self.cboScript.setCurrentText('')            
                self.txtMsg.setText('')
                self.txtMsg_append('找不到對應的測試腳本!',Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        return found


    #按手動修改值
    def on_btnMModify_Click(self):
        subProcName="onManualModifyClick"
        try:     
            if self.txtSpecTRow.toPlainText().strip()=='':
                QMessageBox.information(self,  "手動修正測量值", '請在"指定測試列"內, \n填寫要修改的資料列!',QMessageBox.Ok | QMessageBox.Ok)
                return
            self.txtMsg.setText('')
            self.onlyStopEnable() #設為只有STOP可按  
            self.getSpecTestRow()   #取得指定要測試的ROW
            self.StopTest=False
            self.clearTestJudge()
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        TH=recordHeader
        try:
            for i in self.SpecTRow:
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)
                r=window.TRecord.model.record(i)
                strIName=r.value(TH.col_ItemName)
                odt=r.value(TH.col_TDatetime)
                rComment=f'手動修正測值, 原測值:{r.value(TH.col_Measure_Main)}, 測試時間:{odt}, 備註:{r.value(TH.col_Comment)}'
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                gs="請輸入修正後的主量測值"
                MRdg, okPressed = QInputDialog.getText(self, "手動修正主量測值",gs, QLineEdit.Normal, "")
                MRdg=MRdg.strip()
                if okPressed: 
                    if isnumber(MRdg):
                        meas_main=float(MRdg)                        
                    else:
                        meas_main=MRdg
                else:                    
                    self.StopTest=True
                    break
                r.setValue(TH.col_Comment, rComment)
                r.setValue(TH.col_Measure_Main,meas_main)                

                if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                    r.setValue(TH.col_Judge, 'PASS')                    
                else:
                    r.setValue(TH.col_Judge, 'FAIL')
                self.Update_TRecord(r,True,rStartTime)
                sIndex = self.tbView.model().index(i, 0)
                self.tbView.scrollTo(sIndex) 
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        
        try:
            # sJudge=self.JudgeResult(strIName)
            window.TRecord.model.TestRow=-1
            window.TRecord.tbview.scrollTo(rindex) 
            self.updateTestSummary()
        except Exception as e:
            serr=e
        self.onlyStopDisable() #測試結束, 其他都可按, 只有STOP不能

    def on_btnProTest_Click(self):
        self.btnStop.setEnabled(True)
        self.StopTest=False
        self.UploadTestData()
        self.btnStop.setEnabled(False)
    
    #治具路徑測試    
    def on_btnFixtureTest_Click(self):
        self.StopTest=False
        FPName=''
        self.txtMsg.setText("")        
                        
        self.onlyStopEnable()   #設為只有STOP可按 
        tc=0
        for i in range(0,len(self.cf)):
                if self.cf[i].isChecked():
                    tc+=1
        c=0
        try:
            self.SwitchPath=''
            for i in range(0,len(self.cf)):
                if self.StopTest: break
                if self.cf[i].isChecked():
                    FPName=self.cf[i].text()
                    self.SwitchFix(FPName)
                    c+=1
                    if tc>1 and c<tc:
                        self.delay(5)
            if tc==0:
                self.SwitchFix('Port全不選')
        except Exception as e:
            self.txtMsg_append(str(e))
        self.onlyStopDisable()  #設為只有STOP不可按 

    #清空單筆測試記錄的量測值
    def Blank_TRecord(self, r=QSqlQueryModel.record):
        TH=recordHeader()
        r.setValue(TH.col_Measure_Main,'')
        r.setValue(TH.col_Measure_1,'')
        r.setValue(TH.col_Measure_2,'')
        r.setValue(TH.col_Judge, '')
        r.setValue(TH.col_Comment,'')
        self.Update_TRecord(r,False)
    
    #更新單筆測試記錄
    def Update_TRecord(self, r=QSqlQueryModel.record,WriteLog=True,StartDateTime=''):
        subProcName='Update_TRecord'
        TH=recordHeader()
        try:
            docno=window.TRecord.TestDocNo
            tdatetime = QDateTime.currentDateTime()
            rdt=tdatetime.toString(self.MyDTFormat)
            sSql=f"update TestRecord  set "
            sSql+=f"TDatetime='{rdt}'"

            if r.value(TH.col_Measure_Main)!="":
                sSql+=f", Measure_Main={r.value(TH.col_Measure_Main)}"
                Meas_M=r.value(TH.col_Measure_Main)
            else:
                sSql+=f", Measure_Main=NULL"
                Meas_M='NULL'
            if r.value(TH.col_Deviation_LowLimit)!="":
                sSql+=f", Deviation_LowLimit={r.value(TH.col_Deviation_LowLimit)}"
                DEV_Low_Lim=r.value(TH.col_Deviation_LowLimit)
            else:
                sSql+=f", Deviation_LowLimit=NULL"
                DEV_Low_Lim='NULL'
            if r.value(TH.col_Deviation_UpLimit)!="":
                sSql+=f", Deviation_UpLimit={r.value(TH.col_Deviation_UpLimit)}"
                DEV_Up_Lim=r.value(TH.col_Deviation_UpLimit)
            else:
                sSql+=f", Deviation_UpLimit=NULL"
                DEV_Up_Lim='NULL'    
            if r.value(TH.col_Measure_1)!="":
                sSql+=f", Measure_1={r.value(TH.col_Measure_1)}"
                Meas_1=r.value(TH.col_Measure_1)
            else:
                sSql+=f", Measure_1=NULL"
                Meas_1='NULL'
            if r.value(TH.col_Measure_2)!="":
                sSql+=f", Measure_2={r.value(TH.col_Measure_2)}"
                Meas_2=r.value(TH.col_Measure_2)
            else:
                sSql+=f", Measure_2=NULL"
                Meas_2='NULL'
            if r.value(TH.col_Judge) !="":            
                sSql+=f", Judge='{r.value(TH.col_Judge)}'"
                rJudge=r.value(TH.col_Judge)
            else:
                sSql+=f", Judge=NULL"
                rJudge='NULL'
            if r.value(TH.col_Comment) !="":
                sSql+=f", Comment='{r.value(TH.col_Comment)}'"
                rComment=r.value(TH.col_Comment)
            else:
                sSql+=f", Comment=NULL"
                rComment='NULL'
            sSql+=f", Datetime_Upload=NULL"
            sSql+=f" where TestDocNo='{docno}' and ItemNo={r.value(TH.col_No_1)} and SubItemNo={r.value(TH.col_No_2)}  and TestPointNo={r.value(TH.col_No_3)}"        
            # print(sSql)
            sSql=sSql.replace('\0','')
            AteDB.ConnectSQLite(self.DBPth)
            AteDB.con.execute(sSql)
            AteDB.con.commit()  
            if WriteLog:
                if StartDateTime=='':
                    StartDateTime ='NULL'
                sSql=f"""Insert into TestLog 
                    (
                    TestDocNo, ItemNo, SubItemNo, TestPointNo, 
                    Workstation, Staff, 
                    Measure_Main, Deviation_LowLimit, Deviation_UpLimit, Measure_1, Measure_2, Judge, Comment,
                    Datetime_Start, Datetime_End 
                    ) 
                    VALUES 
                    (
                        '{docno}',{r.value(TH.col_No_1)}, {r.value(TH.col_No_2)} ,{r.value(TH.col_No_3)},
                        '{self.ltxtWorkstation.text()}', '{self.ltxtInspector.text()}', 
                        {Meas_M}, {DEV_Low_Lim}, {DEV_Up_Lim}, {Meas_1}, {Meas_2}, '{rJudge}', '{rComment}', 
                        '{StartDateTime}', '{rdt}'

                    ) """    
                sSql=sSql.replace('\0','')
                sSql=sSql.replace("'NULL'", "NULL") 
                AteDB.con.execute(sSql)
                AteDB.con.commit()
            AteDB.CloseConnection()
            window.TRecord.showTestData()
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)

    #依項次的PASS, FAIL數, 傳回該項的判定
    def JudgeResult(self, ItemName):
        subProcName='JudgeResult'
        TH=recordHeader() 
        CountI=0; CountP=0;  CountF=0
        sJudge=''
        try:
            for i in range(0,window.TRecord.model.rowCount()):            
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),ItemName): continue #不是此測項, 不執行此圈
                js=str(r.value(TH.col_Judge)).upper()
                if js=='PASS': 
                    CountP+=1
                elif js=='FAIL':
                    CountF+=1
                CountI+=1   #計算此測項的測項數量            
            
            if CountF>0:            #FAIL單項次>0, 即判定此項為FAIL
                sJudge='FAIL'
            elif CountP==CountI:    #PASS項次=要測試項次, 才判定PASS
                sJudge='PASS'
            else:
                sJudge='---'        #没有FAIL, 但PASS項次<要測項次, 表示未測完
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        return sJudge

    #取得要設定pin mode out命令字串
    def getDIOPinMode_cmd(self):
        cmdF=''
        try:
            AteDB.ConnectSQLite(self.DBPth)
            sSql=f"select PortNo, Pins from PortAssign_Mega order by PortNo"
            rsPort=AteDB.Query(sSql)        
            AteDB.CloseConnection()
            #self.txtMsg_append(f"送出命令給MEGA 2560=>")
            # self.delay(1)
            cmdF=''
            for i in range(0,len(rsPort)): 
                # cmdF=''
                pin=str(rsPort[i]['Pins']).split(",")
                for j in range(0,len(pin)):                    
                    pin[j]=pin[j].strip()                
                    cmdF+=f"[{pin[j]},out]"
            cmdF='PinDIOMode:'+cmdF
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        return cmdF
    
    #把輸出port的pin設為output 模式
    def setDIOPinMode_out(self):
        subProcName="showPathItem"
        try:
            t1=time()
            if DIO.instrConnected==0:
                return
            if self.model.rowCount()==0:
                return           
             
            cmdF=self.getDIOPinMode_cmd()                   
            DIO.sendCommand(cmdF)
            self.delay(0.05)            
            self.txtMsg_append(f"設定DIO MEGA 2560腳位輸出模式, 花費:{str(round(time()-t1,1))}秒",Qt.blue)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)

    #切換治具
    def SwitchFix(self, SFPath="", Hint=""):     
        subProcName="SwitchFix" 
        
        #t1=time()
        TH=recordHeader()
        s=""
        try:
            DIO.openPort()
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        try:
            for i in range(0,len(self.cf)):
                FPName=self.cf[i].text()
                font=self.cf[i].font()
                if self.cf[i].isChecked() and FPName==SFPath:                    
                    self.cf[i].setStyleSheet('color:green')
                    font.setBold(True)
                    self.cf[i].setFont(font) 
                    rect_cf=self.cf[i].geometry()
                    self.scrollArea_3.ensureVisible(0,rect_cf.y())
                else:
                    self.cf[i].setStyleSheet('color:black')
                    font.setBold(False)
                    self.cf[i].setFont(font)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        QtWidgets.QApplication.processEvents()
        try:
            if self.SwitchPath!=SFPath:
                AteDB.ConnectSQLite(self.DBPth)
                sSql=f"select PortNo, Pins from PortAssign_Mega order by PortNo"
                rsPort=AteDB.Query(sSql)
                sSql=f"select PName, PDesc, Port0, Port1, Port2, Port3, Port4, Port5 from FixturePath where PName='{SFPath}'"
                rsFix=AteDB.Query(sSql)
                AteDB.CloseConnection()

                if self.chkFixManual.isChecked():
                    reply = QMessageBox.information(self, '手動換線', f"請手動換線至\n{SFPath}", 
                            QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
                    if reply == QMessageBox.Cancel:
                        s="Cancel"                
                else:
                    self.txtMsg_append(f"自動換線至{SFPath}...(注意:需使用ATE治具)",Qt.blue)
                    try:
                        i=0
                        fmega=False
                        DIO_IDN=''
                        DIO_IDN=DIO.get_IDN()                        
                        if str(DIO_IDN).lower().find('mega')>=0:
                            fmega=True
                        else:
                            DIO.openPort()
                            # self.delay(1)               
                            for i in range(0,10):
                                self.delay(0.5)
                                DIO_IDN=DIO.get_IDN()
                                if str(DIO_IDN).lower().find('mega')>=0:
                                    fmega=True
                                    break                                
                    except Exception as e:
                        self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red) 
                    if fmega:
                        p=[]
                        p.clear()
                        # self.txtMsg_append(f"治具路徑名稱:{SFPath}")               
                        cmdF=''
                        for i in range(0,len(rsPort)): 
                            pin=str(rsPort[i]['Pins']).split(",")
                            p.append(pin)
                            if len(rsFix)>0:
                                v=rsFix[0]['Port'+str(i)]
                            else:
                                v=''        #没找到符合的治具路徑, 當做都没有PIN輸出HI
                            v=xstr(v)       #V為None時, 會轉為''
                            if v!='':    
                                pv=str(v).split(',')
                                for j in range(0,len(pv)):
                                    pv[j]=str(pv[j]).strip()    #清除前後空格
                                for j in range(0,len(p[i])):
                                    if str(j) in pv:
                                        p[i][j]='['+p[i][j]+',1]'
                                    else:
                                        p[i][j]='['+p[i][j]+',0]'
                                    cmdF+=p[i][j]
                            else:
                                for j in range(0,len(p[i])):                    
                                    p[i][j]='['+p[i][j]+',0]'
                                    cmdF+=p[i][j]
                        for ti in range(0,3):
                            self.delay(0.05)
                            # self.setDIOPinMode_out() #設定MEGA 2560 Pin mode為OUTPUT
                            cmdM=self.getDIOPinMode_cmd()                   
                            DIO.sendCommand(cmdM)   #設定MEGA 2560 Pin mode為OUTPUT
                            self.delay(0.05)
                            dio_rec_cmd1=''
                            dio_rec_cmd2=''
                            dio_rec_cmd1=DIO.sendQuery('lastcmd?')                            
                            # print (f"DIO最後處理的命令:{dio_rec_cmd1}")
                            self.delay(0.05)
                            cmdF="out:"+cmdF
                            # self.txtMsg_append(f"送命令至DIO => {cmdF}",Qt.black)                      
                            DIO.sendCommand(cmdF)  
                            self.delay(0.05)
                            dio_rec_cmd2=DIO.sendQuery('lastcmd?')
                            # print (f"DIO最後處理的命令:{dio_rec_cmd2}") 
                            
                            if dio_rec_cmd1.lower().find(cmdM.lower())>=0 and dio_rec_cmd2.lower().find(cmdF.lower())>=0: #比較送出的命令與MEGA2560收到的命令是否相同
                                self.SwitchPath=SFPath
                                s="AUTO_OK"
                                if ti>0:
                                    self.txtMsg_append(f'第{ti}次發送DIO out OK.',Qt.blue)
                                break
                            else:
                                self.txtMsg_append(f'第{ti}次發送DIO out NG.',Qt.red)
                                s="AUTO_NG"
                    else:
                        self.txtMsg_append(f"{subProcName} Error: 找不到DIO控制器(MEGA 2560)",Qt.red)
                        s="AUTO_NG"
                    DIO.closePort()
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)        
        return s
    

    #PQC測試最外層-----------------------------------------------------------------------------------------------------------------------------
    def ExecPQC(self, ItemName='', FixUse='all'):
        subProcName="ExecPQC"
        print(f"{self.DUT.VisaDescription}")
        if self.DUT is None:
            print("self.DUT是空")
            self.DUT = MyInstr.Instrument()
            print("已建立self.DUT")
        try:
            if self.FirstExecPQC:
                
                print(f"{self.DUT.VisaDescription}")
                print(f"{self.DUT}")
            if self.DUT is None:
                print("self.DUT is None")
            elif not hasattr(self.DUT, "VisaDescription"):
                print("self.DUT 沒有 VisaDescription")
            else:
                print(f"self.DUT.VisaDescription: {self.DUT.VisaDescription}")  
                print(f"執行ExecPQC前: {self.DUT.VisaDescription}")
                #print(f"{self.VisaDescription}")
                
                self.DUT.openPort()
                self.DUT.instrIDN=self.DUT.get_IDN()
                self.DUTIDN=self.DUT.instrIDN   #updateTestSummary會用到此變數
                if self.DUT.instrIDN.find(self.ltxtDUTModel.text())<0 or self.DUT.instrIDN.find(self.ltxtDUTSn.text())<0:
                    font = self.ledDUT_IDN.font()
                    font.setBold(True)
                    font.setPointSize(11)
                    self.ledDUT_IDN.setFont(font)
                    self.ledDUT_IDN.setStyleSheet('color:red')
                else:
                    font = self.ledDUT_IDN.font()
                    font.setBold(False)
                    font.setPointSize(11)
                    self.ledDUT_IDN.setFont(font)
                    self.ledDUT_IDN.setStyleSheet('color:blue')
                self.ledDUT_IDN.setText(self.DUT.instrIDN)
        
                QtWidgets.QApplication.processEvents()
                print(self.DUT.VisaDescription)
                #self.DUT.closePort()
            sJudge='Fail'
            IName=ItemName[ItemName.find('.')+1:]
            if IName in ['IR電壓設定準確度','IR電壓調整率']:
                sJudge=self.PQC_IR_V_Accuracy(IName,FixUse)
            elif IName in ['電壓設定準確度','電壓調整率','電壓量測準確度']:            
                sJudge=self.PQC_HV_Accuracy(IName,FixUse)
            elif IName in ['電流量測準確度']:
                sJudge=self.I_Meas_Accuracy(IName,FixUse)
            elif IName.lower() in ['cut off 電流準確度']:
                sJudge=self.I_CutOff_Accuracy(IName.lower(),FixUse)
            elif IName in ['IR絕緣阻抗量測準確度']:
                sJudge=self.IR_Meas_Accuracy(IName,FixUse)
            elif IName in ['GB電流設定準確度']:
                sJudge=self.GB_I_Set_Accuracy(IName,FixUse)
            elif IName in ['GB電阻量測準確度']:
                sJudge=self.GB_R_Meas_Accuracy(IName,FixUse)
            elif IName in ['空戴漏電流確認','空載漏電流確認']:
                sJudge=self.I_Leakage(FixUse)
            elif IName in ['CONT電阻量測準確度']:
                sJudge=self.CONT_Meas_Accuracy(IName,FixUse)
            elif IName in ['背板輸出確認','掃描功能確認','ARC功能確認','REMOTE功能確認','Signal IO功能確認','GPIB功能確認','Initial']:
                sJudge=self.Manual_Check_Func(IName,FixUse)
            elif IName in ['寫入USB識別碼','寫入USB識別碼確認']:
                sJudge=self.Write_USB_ID_Check(IName,FixUse)
            elif IName in ['寫入機器序號']:
                sJudge=self.Write_DUT_SN(FixUse)
            elif IName in ['寫入機器日期時間']:
                sJudge=self.Write_DUT_DateTime(FixUse)
            elif IName in ['PEL_CC準確度']:
                sJudge=self.PEL_CC_Accuracy(IName,FixUse) 
            elif IName in ['PEL_CC表頭準確度']:
                sJudge=self.PEL_CC_Mea_Readback(IName,FixUse)
            elif IName in ['PEL_CR準確度']:
                sJudge=self.PEL_CR_Accuracy(IName,FixUse)
            elif IName in ['PEL_CV準確度']:
                sJudge=self.PEL_CV_Accuracy(IName,FixUse)
            elif IName in ['PEL_CV表頭準確度']:
                sJudge=self.PEL_CV_Mea_Readback(IName,FixUse)
            elif IName in ['PEL_CP準確度']:
                sJudge=self.PEL_CP_Accuracy(IName,FixUse)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.FirstExecPQC=False          
        return sJudge

    
    #手動填高壓值
    def HVM_Manual(self):
        if not self.StopTest:            
            gvOK=False
            g=0
            gs="請輸入高壓表讀值(KV)"
            while(g<2 and gvOK==False):                         
                shvm, okPressed = QInputDialog.getText(self, "手填高壓表讀值",gs, QLineEdit.Normal, "")
                shvm=shvm.strip()
                if okPressed: 
                    if isnumber(shvm):
                        rdg_hvm=float(shvm)
                        if rdg_hvm<7:
                            gvOK=True
                            break
                        else:
                            gs="值>7KV, 請再確認, 重輸入高壓表讀值(KV)"
                            gvOK=False
                    else:
                        gs="請再確認, 重輸入高壓表讀值(KV)"
                        gvOK=False
                else:
                    rdg_hvm=''
                    self.StopTest=True
                    break
        return rdg_hvm
    #PEL預備測項
    def PEL_CC_Accuracy(self,strIName='',FixUse='all'):
        TH=recordHeader()
        sJudge=''
        try:
            #if self.DUT.instrConnected==0 or not(PSW.instrConnected==1):     #確認此項需要的設備是否有連接
             #   self.txtMsg_append(f"無法測試{strIName},待測機或PSW未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge
            #if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                                    #非手動換線又治具未連線
             #   self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge   
            #self.txtMsg_append(f"*** {strIName} ***",Qt.blue)
            QtWidgets.QApplication.processEvents()
            print(self.DUT.VisaDescription)
            self.DUT.openPort()
            self.DUT.CLS()
            self.DUT.RST()
            i2t=''
            i3t=''
            strM=''
            self.PSC_CCACC = []
            for i in range(0,window.TRecord.model.rowCount()):                  #開始掃報告的每一列
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break                                         #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue       #有指定測試列, 但目前列, 不在指定測試列, 略過
                r=window.TRecord.model.record(i)                                #取得目前指標列的內容
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue     #測試項名稱與此副程式要測試的項目不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix(r.value(TH.col_Cond_Fixture))   #切換治具                
                self.delay(0.5)
                stm=monotonic()                                 #設等待的時間起點
                self.DUT.CLS()
                PCS.RST()
                PSW.RST()
                DMM.RST()
                self.DUT.setOper_Mode("CC")
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('LOW')>=0:
                    self.DUT.setCurrent_Range("LOW")
                    self.DUT.setCC_Current(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(1)##CH####
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('MID')>=0:
                    self.DUT.setCurrent_Range("MID")
                    self.DUT.setCC_Current(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(7.1)##CH####
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('HIGH')>=0:
                    self.DUT.setCurrent_Range("HIGH")
                    self.DUT.setCC_Current(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(20)##CH####
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                MeasDut_List = [
                        f"Oper Mode: {self.DUT.getOper_Mode()}",
                        f"Current Range: {self.DUT.getCurrent_Range()}",
                        f"CC Current: {self.DUT.getCC_Current()}"
                        ]
                MeasDut = "\n".join(MeasDut_List)
                self.txtMsg_append(f"輸出前DUT狀態:\n{str(MeasDut)}")
                self.DUT.setInput_State(1)
                self.txtMsg_append(f"DUT_LOAD_ON...")
                PSW.setOutput_State(1,delay_before=1.5)
                PCS_read=PCS.getReading_DCI(delay_before=2)
                self.PSC_CCACC.append(PCS_read)
                print(self.PSC_CCACC)
                DMM_CC_MEAS_V=DMM.getMainReading(delay_before=1.5)
                PEL_CC_SET_V=self.DUT.getCC_Current()
                self.txtMsg_append(f"電表錶頭值:{str(DMM_CC_MEAS_V)}")##CH####
                self.txtMsg_append(f"PEL電流設定:{str(PEL_CC_SET_V)}")##CH####
                model_config = {
                                'PEL-3021':   {'LOW': 0.35, 'MIDDLE':35, 'HIGH': 35,  'VAR_1': 500000,  'MODE': 'A'},
                                'PEL-3041':   {'LOW': 0.7, 'MIDDLE':70,  'HIGH': 70,  'VAR_1': 500000,  'MODE': 'A'},
                                'PEL-3111':   {'LOW': 2.1, 'MIDDLE':210, 'HIGH': 210, 'VAR_1': 500000,  'MODE': 'A'},
                                'PEL-3021H':  {'LOW': 87.5,'MIDDLE':8.75,'HIGH': 8.75, 'VAR_1': 3240000, 'VAR_2': 2210000,'MODE': 'B'},
                                'PEL-3041H':  {'LOW': 175, 'MIDDLE':17.5, 'HIGH': 17.5,  'VAR_1': 2210000, 'MODE': 'C'},
                                'PEL-3111H':  {'LOW': 0.525,'MIDDLE':52.5,'HIGH':52.5,  'VAR_1': 2210000, 'MODE': 'D'},
                                'PEL-3031AE': {'LOW': 60,'MIDDLE':60,     'HIGH':60,  'VAR_1': 500000,  'MODE': 'D'},
                                'PEL-3032AE': {'LOW': 15,'MIDDLE':15,     'HIGH': 15,  'VAR_1': 500000,  'MODE': 'D'},
                                }

                DUT_CP_MODE = self.DUT.get_IDN().split(",")[1].strip() ##CH###             
                oper_mode = self.DUT.getCurrent_Range()

                
                
                if DUT_CP_MODE in model_config:
                    cfg = model_config[DUT_CP_MODE]
                    fs = cfg.get(oper_mode)
                    VAR_1 = cfg.get("VAR_1")
                    VAR_2 = cfg.get("VAR_2")
                    mode = cfg.get("MODE")

                    if mode == 'A':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) - (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    elif mode == 'B' and oper_mode == 'HIGH':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    elif mode == 'B' and oper_mode == 'MIDDLE':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_2)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_2)
                    elif mode == 'B' and oper_mode == 'LOW':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1 *1000)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1 *1000)
                    elif mode == 'C' and oper_mode == 'HIGH':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    elif mode == 'C' and oper_mode == 'MIDDLE':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    elif mode == 'C' and oper_mode == 'LOW':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1 *1000)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1 *1000)
                    elif mode == 'D' and oper_mode == 'HIGH':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    elif mode == 'D' and oper_mode == 'MIDDLE':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    elif mode == 'D' and oper_mode == 'LOW':
                        CL = PEL_CC_SET_V - ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                        CU = PEL_CC_SET_V + ((0.002 * PEL_CC_SET_V) + (0.001 * fs)) + (DMM_CC_MEAS_V / VAR_1)
                    
                    

                if strIName=='PEL_CC準確度':
                    r.setValue(TH.col_Measure_Main,PCS_read)
                    r.setValue(TH.col_LowLimit,CL)
                    r.setValue(TH.col_UpLimit,CU)
                    r.setValue(TH.col_Measure_1,PEL_CC_SET_V)
                    r.setValue(TH.col_Measure_2,DMM_CC_MEAS_V)
                    if  r.value(TH.col_Measure_Main)!='':
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                        if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                        r.setValue(TH.col_Judge, 'FAIL')
                        rComment+=f",Status: {MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    self.Update_TRecord(r,True,rStartTime)
                    PSW.setOutput_State(0,delay_after=1)
                    self.DUT.CLS()
                    self.DUT.setInput_State(0,delay_before=0.5)            #輸出OFF
                    self.txtMsg_append(f"DUT_LOAD_OFF.")
                    self.delay(0.5)
                else:
                    self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
                
                #self.txtMsg_append(f"輸出後DUT狀態:\n{str(MeasDut)}\nSYS Err:{DErr}")
                #self.txtMsg_append(f"Measure?傳回:\n{self.DUT.getMeasure_RawStr()}")    
                 
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        PSW.setOutput_State(0,delay_after=1.5)
        self.DUT.CLS()
        self.DUT.setInput_State(0)            #輸出OFF
        self.DUT.closePort()
        PSW.closePort()
        sJudge=self.JudgeResult(strIName)   #判斷此測試PASS或FAIL
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge
        
    def PEL_CC_Mea_Readback(self,strIName='',FixUse='all'):
        TH=recordHeader()
        sJudge=''
        try:
            #if self.DUT.instrConnected==0 or not(PSW.instrConnected==1):     #確認此項需要的設備是否有連接
             #   self.txtMsg_append(f"無法測試{strIName},待測機或PSW未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge
            #if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                                    #非手動換線又治具未連線
             #   self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge   
            #self.txtMsg_append(f"*** {strIName} ***",Qt.blue)
            QtWidgets.QApplication.processEvents()
            print(self.DUT.VisaDescription)
            self.DUT.openPort()
            self.DUT.CLS()
            self.DUT.RST()
            i2t=''
            i3t=''
            strM=''
            PCS_CCMEAS_READBK=''
            PEL_MEAS_READBK=''
            AL=''
            AU=''
            for i in range(0,window.TRecord.model.rowCount()):                  #開始掃報告的每一列
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break                                         #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue       #有指定測試列, 但目前列, 不在指定測試列, 略過
                r=window.TRecord.model.record(i)                                #取得目前指標列的內容
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue     #測試項名稱與此副程式要測試的項目不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix(r.value(TH.col_Cond_Fixture))   #切換治具                
                self.delay(0.5)
                stm=monotonic()                                 #設等待的時間起點
                self.DUT.CLS()
                PCS.RST()
                PSW.RST()
                DMM.RST()
                self.DUT.setOper_Mode("CC")
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('LOW')>=0:
                    self.DUT.setCurrent_Range("LOW")
                    self.DUT.setCC_Current(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(1)
                    psw_vol=PSW.getVoltage()
                    self.txtMsg_append(f"PSW設定電壓值:{str(psw_vol)}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('MID')>=0:
                    self.DUT.setCurrent_Range("MID")
                    self.DUT.setCC_Current(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(7.1)
                    psw_vol=PSW.getVoltage()
                    self.txtMsg_append(f"PSW設定電壓值:{str(psw_vol)}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('HIGH')>=0:
                    self.DUT.setCurrent_Range("HIGH")
                    self.DUT.setCC_Current(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(20)
                    psw_vol=PSW.getVoltage()
                    self.txtMsg_append(f"PSW設定電壓值:{str(psw_vol)}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                MeasDut_List = [
                        f"Oper Mode: {self.DUT.getOper_Mode()}",
                        f"Current Range: {self.DUT.getCurrent_Range()}",
                        f"CC Current: {self.DUT.getCC_Current()}"
                        ]
                MeasDut = "\n".join(MeasDut_List)
                self.txtMsg_append(f"輸出前DUT狀態:\n{str(MeasDut)}")
                self.DUT.setInput_State(1)
                self.txtMsg_append(f"DUT_LOAD_ON...")
                PSW.setOutput_State(1,delay_before=1.5)
                PCS_CCMEAS_READBK=PCS.getReading_DCI(delay_before=1.5)
                PEL_MEAS_READBK=self.DUT.getMeasure_Current(delay_before=1.5)
                PEL_CC_SET_V=self.DUT.getCC_Current()
                model_config = {
                                'PEL-3021':   {'LOW': 0.35, 'MIDDLE':35, 'HIGH': 35,   'MODE': 'A'},
                                'PEL-3041':   {'LOW': 0.7, 'MIDDLE':70,  'HIGH': 70,   'MODE': 'A'},
                                'PEL-3111':   {'LOW': 2.1, 'MIDDLE':210, 'HIGH': 210,  'MODE': 'A'},
                                'PEL-3021H':  {'LOW': 87.5,'MIDDLE':8.75,'HIGH': 8.75, 'MODE': 'A'},
                                'PEL-3041H':  {'LOW': 175, 'MIDDLE':17.5, 'HIGH': 17.5, 'MODE': 'A'},
                                'PEL-3111H':  {'LOW': 0.525,'MIDDLE':52.5,'HIGH':52.5, 'MODE': 'A'},
                                'PEL-3031AE': {'LOW': 60,   'MIDDLE':60,     'HIGH':60,   'MODE': 'B'},
                                'PEL-3032AE': {'LOW': 15,   'MIDDLE':15,     'HIGH': 15,  'MODE': 'B'},
                                }

                DUT_CP_MODE = self.DUT.get_IDN().split(",")[1].strip() ##CH###              
                oper_mode = self.DUT.getCurrent_Range()       

                if DUT_CP_MODE in model_config:
                    cfg = model_config[DUT_CP_MODE]
                    fs = cfg.get(oper_mode)
                    mode = cfg.get("MODE")
                    if mode == 'A':
                        AL= float(PCS_CCMEAS_READBK - ((0.2 / 100) * PCS_CCMEAS_READBK + (0.3 / 100) * fs))
                        print(f"計算結果: {AL}")
                        AU= float(PCS_CCMEAS_READBK + ((0.2 / 100) * PCS_CCMEAS_READBK + (0.3 / 100) * fs))
                        print(f"計算結果: {AU}")
                    if mode == 'B'and oper_mode == 'LOW':
                        AL= float(PCS_CCMEAS_READBK - ((0.1 / 100) * PCS_CCMEAS_READBK + (0.1 / 100) * fs))
                        print(f"計算結果: {AL}")
                        AU= float(PCS_CCMEAS_READBK + ((0.1 / 100) * PCS_CCMEAS_READBK + (0.1 / 100) * fs))
                        print(f"計算結果: {AU}")
                    if mode == 'B'and oper_mode == 'HIGH':
                        AL= float(PCS_CCMEAS_READBK - ((0.1 / 100) * PCS_CCMEAS_READBK + (0.2 / 100) * fs))
                        print(f"計算結果: {AL}")
                        AU= float(PCS_CCMEAS_READBK + ((0.1 / 100) * PCS_CCMEAS_READBK + (0.2 / 100) * fs))
                        print(f"計算結果: {AU}")

                    
                if strIName=='PEL_CC表頭準確度':
                    r.setValue(TH.col_Deviation_LowLimit,AL)
                    r.setValue(TH.col_Deviation_UpLimit,AU)
                    r.setValue(TH.col_Measure_Main,PEL_MEAS_READBK)
                    r.setValue(TH.col_Measure_1,PEL_CC_SET_V)
                    r.setValue(TH.col_Measure_2,PCS_CCMEAS_READBK)
                    if  r.value(TH.col_Measure_Main)!='':
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                        if r.value(TH.col_Measure_Main) > r.value(TH.col_Deviation_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_Deviation_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                        r.setValue(TH.col_Judge, 'FAIL')
                        rComment+=f",Status: {MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    self.Update_TRecord(r,True,rStartTime)
                    PSW.setOutput_State(0)
                    self.DUT.CLS()
                    self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
                    self.txtMsg_append(f"DUT_LOAD_OFF.")
                    self.delay(0.5)
                else:
                    self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
                
                #self.txtMsg_append(f"輸出後DUT狀態:\n{str(MeasDut)}\nSYS Err:{DErr}")
                #self.txtMsg_append(f"Measure?傳回:\n{self.DUT.getMeasure_RawStr()}")    
                 
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        PSW.setOutput_State(0)
        self.DUT.CLS()
        self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
        self.DUT.closePort()
        PSW.closePort()
        sJudge=self.JudgeResult(strIName)   #判斷此測試PASS或FAIL
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge

    def PEL_CR_Accuracy(self,strIName='',FixUse='all'):
        TH=recordHeader()
        sJudge=''
        try:
            #if self.DUT.instrConnected==0 or not(PSW.instrConnected==1):     #確認此項需要的設備是否有連接
             #   self.txtMsg_append(f"無法測試{strIName},待測機或PSW未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge
            #if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                                    #非手動換線又治具未連線
             #   self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge   
            #self.txtMsg_append(f"*** {strIName} ***",Qt.blue)
            QtWidgets.QApplication.processEvents()
            print(self.DUT.VisaDescription)
            self.DUT.openPort()
            self.DUT.CLS()
            self.DUT.RST()
            i2t=''
            i3t=''
            strM=''
            for i in range(0,window.TRecord.model.rowCount()):                  #開始掃報告的每一列
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break                                         #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue       #有指定測試列, 但目前列, 不在指定測試列, 略過
                r=window.TRecord.model.record(i)                                #取得目前指標列的內容
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue     #測試項名稱與此副程式要測試的項目不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix(r.value(TH.col_Cond_Fixture))   #切換治具                
                self.delay(0.5)
                stm=monotonic()                                 #設等待的時間起點
                self.DUT.CLS()
                PCS.RST()
                PSW.RST()
                DMM.RST()
                DUT_CR_MODE = self.DUT.get_IDN()
                self.DUT.setOper_Mode("CR")
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('LOW')>=0:
                    self.DUT.setCurrent_Range("LOW")
                    self.DUT.setCR_Resistance(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(1)
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('MID')>=0:
                    self.DUT.setCurrent_Range("MID")
                    self.DUT.setCR_Resistance(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(7.1)
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('HIGH')>=0:
                    self.DUT.setCurrent_Range("HIGH")
                    self.DUT.setCR_Resistance(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    PSW.setCurrent(21.6)
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                MeasDut_List = [
                        f"Oper Mode: {self.DUT.getOper_Mode()}",
                        f"Current Range: {self.DUT.getCurrent_Range()}",
                        f"CR Resistance: {self.DUT.getCR_Resistance()}"
                        ]
                MeasDut = "\n".join(MeasDut_List)
                self.txtMsg_append(f"輸出前DUT狀態:\n{str(MeasDut)}")
                self.DUT.setInput_State(1)
                self.txtMsg_append(f"DUT_LOAD_ON...")
                PSW.setOutput_State(1,delay_before=1.5)
                PCS_CRMEAS=PCS.getReading_DCI(delay_before=1.5)
                DMM_CR_MEAS_V=DMM.getMainReading(delay_before=1.5)
                PEL_CR_SET = self.DUT.getCR_Resistance()
                
                # FS 參數表
                FS_TABLE = {
                    "PEL-3021": {"LOW": 0.35, "HIGH": 35, "V_DEN": 500000},
                    "PEL-3021H": {"LOW": 87.5, "HIGH": 8.75, "V_DEN": 3240000},
                    "PEL-3041": {"LOW": 0.7, "HIGH": 70, "V_DEN": 500000},
                    "PEL-3041H": {"LOW": 175, "HIGH": 17.5, "V_DEN": 2210000},
                    "PEL-3111": {"LOW": 2.1, "HIGH": 210, "V_DEN": 500000},
                    "PEL-3111H": {"LOW": 0.525, "HIGH": 52.5, "V_DEN": 2210000},
                    "PEL-3031E": {"LOW": 60, "HIGH": 60, "V_DEN": 0},
                    "PEL-3032E": {"LOW": 6, "HIGH": 6 , "V_DEN": 0},
                            }

                DUT_CR_MODE = {
                    "IDN": self.DUT.get_IDN(),##.split(",")[1].strip(),
                    "Oper_Mode": self.DUT.getCurrent_Range()
                                }

                # 轉換 OPER_MODE，將 MID 視為 HIGH
                oper_mode = DUT_CR_MODE["Oper_Mode"]
                if oper_mode == "MIDDLE":
                    oper_mode = "HIGH"

                model = next((m for m in FS_TABLE if m in DUT_CR_MODE["IDN"]), None)

                PEL_CR_RE = round((1 / PEL_CR_SET) * 1000, 4)

                if model and oper_mode in FS_TABLE[model]:
                    FS = FS_TABLE[model][oper_mode]
                    V_DEN = FS_TABLE[model]["V_DEN"]

                    
                    # LOW 模式使用不同標準計算
                    if oper_mode == "LOW":
                        if 'PEL-3111H' in model:
                            PEL_CR_STANDARD = DMM_CR_MEAS_V * PEL_CR_RE / 1000
                            RL = PEL_CR_STANDARD - ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN)
                            RU = PEL_CR_STANDARD + ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN)
                        elif model.endswith("H"):
                            PEL_CR_STANDARD = DMM_CR_MEAS_V * PEL_CR_RE
                            RL = PEL_CR_STANDARD - ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN * 1000)
                            RU = PEL_CR_STANDARD + ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN * 1000)
                        elif model.endswith("E"):
                            value = PEL_CR_RE - ((0.003 * PEL_CR_RE) + (0.01 * FS * 1000)) + (1 / 500000 * 1000)
                            RL= max(value, 0)
                            RU = PEL_CR_RE + ((0.003 * PEL_CR_RE) + (0.01 * FS * 1000)) + (1 / 500000 * 1000)
                        else:
                            PEL_CR_STANDARD = DMM_CR_MEAS_V * PEL_CR_RE / 1000
                            RL = PEL_CR_STANDARD - ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) - (DMM_CR_MEAS_V / V_DEN)
                            RU = PEL_CR_STANDARD + ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN)
                    elif model in ['PEL-3021H', 'PEL-3041H']:
                        PEL_CR_STANDARD = DMM_CR_MEAS_V * PEL_CR_RE / 1000
                        RL = PEL_CR_STANDARD - ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN)
                        RU = PEL_CR_STANDARD + ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN)
                    elif model.endswith("E"):
                            value = PEL_CR_RE - ((0.003 * PEL_CR_RE) + (0.01 * FS * 1000)) + (1 / 500000 * 1000)
                            RL= max(value, 0)
                            RU = PEL_CR_RE + ((0.003 * PEL_CR_RE) + (0.01 * FS * 1000)) + (1 / 500000 * 1000)
                    else:  # HIGH & MID
                        PEL_CR_STANDARD = DMM_CR_MEAS_V * PEL_CR_RE / 1000
                        RL = PEL_CR_STANDARD - ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) - (DMM_CR_MEAS_V / V_DEN)
                        RU = PEL_CR_STANDARD + ((0.005 * PEL_CR_STANDARD) + (0.005 * FS)) + (DMM_CR_MEAS_V / V_DEN)
                else:
                    print("找不到對應的型號或操作模式")

                if strIName=='PEL_CR準確度':
                    if model.endswith("E"):
                        PCS_CRMEAS_mS = (PCS_CRMEAS / DMM_CR_MEAS_V)*1000
                        r.setValue(TH.col_Measure_Main,PCS_CRMEAS_mS)
                    else:    
                        r.setValue(TH.col_Measure_Main,PCS_CRMEAS)
                    r.setValue(TH.col_Measure_1,DMM_CR_MEAS_V)
                    r.setValue(TH.col_LowLimit,RL)
                    r.setValue(TH.col_UpLimit,RU)
                    if  r.value(TH.col_Measure_Main)!='':
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                        if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                        r.setValue(TH.col_Judge, 'FAIL')
                        rComment+=f",Status: {MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    self.Update_TRecord(r,True,rStartTime)
                    PSW.setOutput_State(0)
                    self.DUT.CLS()
                    self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
                    self.txtMsg_append(f"DUT_LOAD_OFF.")
                    self.delay(0.5)
                else:
                    self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
                
                #self.txtMsg_append(f"輸出後DUT狀態:\n{str(MeasDut)}\nSYS Err:{DErr}")
                #self.txtMsg_append(f"Measure?傳回:\n{self.DUT.getMeasure_RawStr()}")    
                 
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        PSW.setOutput_State(0)
        self.DUT.CLS()
        self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
        self.DUT.closePort()
        PSW.closePort()
        sJudge=self.JudgeResult(strIName)   #判斷此測試PASS或FAIL
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge
    
    def PEL_CV_Accuracy(self,strIName='',FixUse='all'):
        TH=recordHeader()
        sJudge=''
        try:
            #if self.DUT.instrConnected==0 or not(PSW.instrConnected==1):     #確認此項需要的設備是否有連接
             #   self.txtMsg_append(f"無法測試{strIName},待測機或PSW未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge
            #if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                                    #非手動換線又治具未連線
             #   self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge   
            #self.txtMsg_append(f"*** {strIName} ***",Qt.blue)
            QtWidgets.QApplication.processEvents()
            print(self.DUT.VisaDescription)
            self.DUT.openPort()
            self.DUT.CLS()
            self.DUT.RST()
            i2t=''
            i3t=''
            strM=''
            for i in range(0,window.TRecord.model.rowCount()):                  #開始掃報告的每一列
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break                                         #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue       #有指定測試列, 但目前列, 不在指定測試列, 略過
                r=window.TRecord.model.record(i)                                #取得目前指標列的內容
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue     #測試項名稱與此副程式要測試的項目不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix(r.value(TH.col_Cond_Fixture))   #切換治具                
                self.delay(0.5)
                stm=monotonic()                                 #設等待的時間起點
                self.DUT.CLS()
                PCS.RST()
                PSW.RST()
                DMM.RST()
                self.DUT.setOper_Mode("CV")
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('LOW')>=0:
                    self.DUT.setCurrent_Range("LOW")
                    self.DUT.setCV_Voltage(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    PSW.setCurrent(1)
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('HIGH')>=0:
                    self.DUT.setCurrent_Range("HIGH")
                    self.DUT.setCV_Voltage(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    PSW.setCurrent(1)
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                MeasDut_List = [
                        f"Oper Mode: {self.DUT.getOper_Mode()}",
                        f"Current Range: {self.DUT.getCurrent_Range()}",
                        f"CV Voltage: {self.DUT.getCV_Voltage()}"
                        ]
                MeasDut = "\n".join(MeasDut_List)
                self.txtMsg_append(f"輸出前DUT狀態:\n{str(MeasDut)}")
                self.DUT.setInput_State(1)
                self.txtMsg_append(f"DUT_LOAD_ON...")
                PSW.setOutput_State(1,delay_before=1.5)
                DMM_CV_MEAS=DMM.getMainReading(delay_before=1.5)
                PEL_CV_SET_V = self.DUT.getCV_Voltage()
                model_config = {
                                'PEL-3021':   {'LOW': 15, 'MID':35, 'HIGH': 150,    'MODE': 'A'},
                                'PEL-3041':   {'LOW': 15, 'MID':70,  'HIGH': 150,   'MODE': 'A'},
                                'PEL-3111':   {'LOW': 15, 'MID':210, 'HIGH': 150,   'MODE': 'A'},
                                'PEL-3021H':  {'LOW': 80,'MID':8.75,    'HIGH': 800,    'MODE': 'A'},
                                'PEL-3041H':  {'LOW': 80, 'MID':17.5,   'HIGH':800,   'MODE': 'A'},
                                'PEL-3111H':  {'LOW': 80,'MID':52.5,    'HIGH':800, 'MODE': 'A'},
                                'PEL-3031AE': {'LOW': 15,'MID':60,     'HIGH':150,  'MODE': 'A'},
                                'PEL-3032AE': {'LOW': 50,'MID':15,     'HIGH':500,  'MODE': 'A'},
                                }

                DUT_CP_MODE = self.DUT.get_IDN().split(",")[1].strip()              
                oper_mode = self.DUT.getCurrent_Range()       

                if DUT_CP_MODE in model_config:
                    cfg = model_config[DUT_CP_MODE]
                    fs = cfg.get(oper_mode)
                    mode = cfg.get("MODE")

                    if mode == 'A':
                        CVL = PEL_CV_SET_V - ((0.001 * PEL_CV_SET_V) + (0.001 * fs))
                        CVU = PEL_CV_SET_V + ((0.001 * PEL_CV_SET_V) + (0.001 * fs))


                if strIName=='PEL_CV準確度':
                    r.setValue(TH.col_Measure_Main,DMM_CV_MEAS)
                    r.setValue(TH.col_LowLimit,CVL)
                    r.setValue(TH.col_UpLimit,CVU)
                    if  r.value(TH.col_Measure_Main)!='':
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                        if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                        r.setValue(TH.col_Judge, 'FAIL')
                        rComment+=f",Status: {MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    self.Update_TRecord(r,True,rStartTime)
                    PSW.setOutput_State(0)
                    self.DUT.CLS()
                    self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
                    self.txtMsg_append(f"DUT_LOAD_OFF.")
                    self.delay(0.5)
                else:
                    self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
                
                #self.txtMsg_append(f"輸出後DUT狀態:\n{str(MeasDut)}\nSYS Err:{DErr}")
                #self.txtMsg_append(f"Measure?傳回:\n{self.DUT.getMeasure_RawStr()}")    
                 
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        PSW.setOutput_State(0)
        self.DUT.CLS()
        self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
        self.DUT.closePort()
        PSW.closePort()
        sJudge=self.JudgeResult(strIName)   #判斷此測試PASS或FAIL
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge
    
    def PEL_CV_Mea_Readback(self,strIName='',FixUse='all'):
        TH=recordHeader()
        sJudge=''
        try:
            #if self.DUT.instrConnected==0 or not(PSW.instrConnected==1):     #確認此項需要的設備是否有連接
             #   self.txtMsg_append(f"無法測試{strIName},待測機或PSW未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge
            #if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                                    #非手動換線又治具未連線
             #   self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge   
            #self.txtMsg_append(f"*** {strIName} ***",Qt.blue)
            QtWidgets.QApplication.processEvents()
            print(self.DUT.VisaDescription)
            self.DUT.openPort()
            self.DUT.CLS()
            self.DUT.RST()
            i2t=''
            i3t=''
            strM=''
            PSW_CVMEAS_READBK=''
            PEL_MEAS_READBK=''
            AL=''
            AU=''
            for i in range(0,window.TRecord.model.rowCount()):                  #開始掃報告的每一列
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break                                         #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue       #有指定測試列, 但目前列, 不在指定測試列, 略過
                r=window.TRecord.model.record(i)                                #取得目前指標列的內容
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue     #測試項名稱與此副程式要測試的項目不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix(r.value(TH.col_Cond_Fixture))   #切換治具                
                self.delay(0.5)
                stm=monotonic()                                 #設等待的時間起點
                self.DUT.CLS()
                PCS.RST()
                PSW.RST()
                DMM.RST()
                self.DUT.setOper_Mode("CV")
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('LOW')>=0:
                    self.DUT.setCurrent_Range("LOW")
                    self.DUT.setCV_Voltage(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    PSW.setCurrent(1)
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('HIGH')>=0:
                    self.DUT.setCurrent_Range("HIGH")
                    self.DUT.setCV_Voltage(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    PSW.setCurrent(1)
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                MeasDut_List = [
                        f"Oper Mode: {self.DUT.getOper_Mode()}",
                        f"Current Range: {self.DUT.getCurrent_Range()}",
                        f"CV Voltage: {self.DUT.getCV_Voltage()}"
                        ]
                MeasDut = "\n".join(MeasDut_List)
                self.txtMsg_append(f"輸出前DUT狀態:\n{str(MeasDut)}")
                self.DUT.setInput_State(1)
                self.txtMsg_append(f"DUT_LOAD_ON...")
                PSW.setOutput_State(1,delay_before=1.5)
                DMM_CVMEAS_READBK=DMM.getMainReading(delay_before=1.5)
                PEL_MEAS_READBK=self.DUT.getMeasure_Voltage(delay_before=1.5)
                
                model_config = {
                                'PEL-3021':   {'LOW': 15, 'MID':35, 'HIGH': 150,    'MODE': 'A'},
                                'PEL-3041':   {'LOW': 15, 'MID':70,  'HIGH': 150,   'MODE': 'A'},
                                'PEL-3111':   {'LOW': 15, 'MID':210, 'HIGH': 150,   'MODE': 'A'},
                                'PEL-3021H':  {'LOW': 80,'MID':8.75,    'HIGH': 800,    'MODE': 'A'},
                                'PEL-3041H':  {'LOW': 80, 'MID':17.5,   'HIGH':800,   'MODE': 'A'},
                                'PEL-3111H':  {'LOW': 80,'MID':52.5,    'HIGH':800, 'MODE': 'A'},
                                'PEL-3031AE': {'LOW': 15,'MID':60,     'HIGH':150,  'MODE': 'A'},
                                'PEL-3032AE': {'LOW': 50,'MID':15,     'HIGH':500,  'MODE': 'A'},
                                }

                DUT_CP_MODE = self.DUT.get_IDN().split(",")[1].strip()              
                oper_mode = self.DUT.getCurrent_Range()       

                if DUT_CP_MODE in model_config:
                    cfg = model_config[DUT_CP_MODE]
                    fs = cfg.get(oper_mode)
                    mode = cfg.get("MODE")

                    if mode == 'A':
                        CVL =DMM_CVMEAS_READBK - ((0.001 *DMM_CVMEAS_READBK) + (0.001 * fs))
                        CVU =DMM_CVMEAS_READBK + ((0.001 *DMM_CVMEAS_READBK) + (0.001 * fs))
                
                print(f"計算結果: {CVL}")
                print(f"計算結果: {CVU}")
                if strIName=='PEL_CV表頭準確度':
                    r.setValue(TH.col_Deviation_LowLimit,CVL)
                    r.setValue(TH.col_Deviation_UpLimit,CVU)
                    r.setValue(TH.col_Measure_Main,PEL_MEAS_READBK)
                    r.setValue(TH.col_Measure_1,DMM_CVMEAS_READBK)
                    if  r.value(TH.col_Measure_Main)!='':
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                        if r.value(TH.col_Measure_Main) > r.value(TH.col_Deviation_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_Deviation_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                        r.setValue(TH.col_Judge, 'FAIL')
                        rComment+=f",Status: {MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    self.Update_TRecord(r,True,rStartTime)
                    PSW.setOutput_State(0)
                    self.DUT.CLS()
                    self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
                    self.txtMsg_append(f"DUT_LOAD_OFF.")
                    self.delay(0.5)
                else:
                    self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
                
                #self.txtMsg_append(f"輸出後DUT狀態:\n{str(MeasDut)}\nSYS Err:{DErr}")
                #self.txtMsg_append(f"Measure?傳回:\n{self.DUT.getMeasure_RawStr()}")    
                 
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        PSW.setOutput_State(0)
        self.DUT.CLS()
        self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
        self.DUT.closePort()
        PSW.closePort()
        sJudge=self.JudgeResult(strIName)   #判斷此測試PASS或FAIL
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge

    def PEL_CP_Accuracy(self,strIName='',FixUse='all'):
        TH=recordHeader()
        sJudge=''
        try:
            #if self.DUT.instrConnected==0 or not(PSW.instrConnected==1):     #確認此項需要的設備是否有連接
             #   self.txtMsg_append(f"無法測試{strIName},待測機或PSW未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge
            #if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                                    #非手動換線又治具未連線
             #   self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
              #  self.delay(0.5)
               # return sJudge   
            #self.txtMsg_append(f"*** {strIName} ***",Qt.blue)
            QtWidgets.QApplication.processEvents()
            print(self.DUT.VisaDescription)
            self.DUT.openPort()
            self.DUT.CLS()
            self.DUT.RST()
            i2t=''
            i3t=''
            strM=''
            for i in range(0,window.TRecord.model.rowCount()):                  #開始掃報告的每一列
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break                                         #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue       #有指定測試列, 但目前列, 不在指定測試列, 略過
                r=window.TRecord.model.record(i)                                #取得目前指標列的內容
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue     #測試項名稱與此副程式要測試的項目不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix(r.value(TH.col_Cond_Fixture))   #切換治具                
                self.delay(0.5)
                stm=monotonic()                                 #設等待的時間起點
                self.DUT.CLS()
                PCS.RST()
                PSW.RST()
                DUT_CP_MODE=self.DUT.get_IDN()
                self.DUT.setOper_Mode("CP")
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('LOW')>=0:
                    self.DUT.setCurrent_Range("LOW")
                    self.DUT.setCP_Power(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                elif strM.find('MID')>=0:
                    self.DUT.setCurrent_Range("MID")
                    self.DUT.setCP_Power(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                    
                elif strM.find('HIGH')>=0:
                    self.DUT.setCurrent_Range("HIGH")
                    self.DUT.setCP_Power(r.value(TH.col_Cond_1))
                    PSW.setVoltage(r.value(TH.col_Cond_2))
                    psw_vol=PSW.getVoltage()
                    print(f"{psw_vol}")
                    PCS.setDCI(r.value(TH.col_Cond_3))
                
                MeasDut_List = [
                        f"Oper Mode: {self.DUT.getOper_Mode()}",
                        f"Current Range: {self.DUT.getCurrent_Range()}",
                        f"CP Power: {self.DUT.getCP_Power()}"
                        ]
                MeasDut = "\n".join(MeasDut_List)
                self.txtMsg_append(f"輸出前DUT狀態:\n{str(MeasDut)}")
                self.DUT.setInput_State(1)
                self.txtMsg_append(f"DUT_LOAD_ON...")
                PSW.setOutput_State(1,delay_before=1.5)
                PEL_CP_SET_POWER=self.DUT.getCP_Power()
                PSW_CP_SET_V=PSW.getVoltage()
                PCS_CP_MEAS_V=PCS.getReading_DCV(delay_before=1.5)
                PCS_CP_MEAS_A=PCS.getReading_DCI(delay_before=1.5)
                PEL_CP_MEAS=PCS_CP_MEAS_A*PCS_CP_MEAS_V
                
                model_config = {
                                'PEL-3021':   {'fs': 1.75,  'VAR': 500000},
                                'PEL-3041':   {'fs': 3.5,   'VAR': 500000},
                                'PEL-3111':   {'fs': 10.5,  'VAR': 500000},
                                'PEL-3021H':  {'fs': 1.75,  'VAR': 3240000},
                                'PEL-3041H':  {'fs': 3.5,   'VAR': 2210000},
                                'PEL-3111H':  {'fs': 10.5,  'VAR': 2210000},
                                'PEL-3031AE': {'fs': 300,   'VAR': 500000},
                                'PEL-3032AE': {'fs': 300,   'VAR': 500000},
                               }
                fs = 3.5
                VAR = 500000
                for model_str, config in model_config.items():
                    if model_str in DUT_CP_MODE:
                        fs = config.get('fs', fs)
                        VAR = config.get('VAR', VAR)
                        break  

                if ['PEL-3041','PEL3021','PEL-3111'] in DUT_CP_MODE:
                    PL= PEL_CP_SET_POWER - ((0.6 / 100) * PEL_CP_SET_POWER + (1.4 / 100) * fs) - (PCS_CP_MEAS_V * PCS_CP_MEAS_V / VAR)
                    PU= PEL_CP_SET_POWER + ((0.6 / 100) * PEL_CP_SET_POWER + (1.4 / 100) * fs) + (PCS_CP_MEAS_V * PCS_CP_MEAS_V / VAR)

                elif ['PEL-3031AE','PEL-3032AE','PEL-3021H','PEL-3041H','PEL-3111H'] in DUT_CP_MODE:
                    PL= PEL_CP_SET_POWER - ((0.6 / 100) * PEL_CP_SET_POWER + (1.4 / 100) * fs) + (PCS_CP_MEAS_V * PCS_CP_MEAS_V / VAR)
                    PU= PEL_CP_SET_POWER + ((0.6 / 100) * PEL_CP_SET_POWER + (1.4 / 100) * fs) + (PCS_CP_MEAS_V * PCS_CP_MEAS_V / VAR)

                print(f"計算結果 PL: {PL}")
                print(f"計算結果 PU: {PU}")

                if strIName=='PEL_CP準確度':
                    r.setValue(TH.col_Measure_Main,PEL_CP_MEAS)
                    r.setValue(TH.col_LowLimit,PL)
                    r.setValue(TH.col_UpLimit,PU)
                    if  r.value(TH.col_Measure_Main)!='':
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                        if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        #self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                        r.setValue(TH.col_Judge, 'FAIL')
                        rComment+=f",Status: {MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    self.Update_TRecord(r,True,rStartTime)
                    PSW.setOutput_State(0)
                    self.DUT.CLS()
                    self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
                    self.txtMsg_append(f"DUT_LOAD_OFF.")
                    self.delay(0.5)
                else:
                    self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
                
                #self.txtMsg_append(f"輸出後DUT狀態:\n{str(MeasDut)}\nSYS Err:{DErr}")
                #self.txtMsg_append(f"Measure?傳回:\n{self.DUT.getMeasure_RawStr()}")    
                 
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: \n{str(e)}",Qt.red)
        PSW.setOutput_State(0)
        self.DUT.CLS()
        self.DUT.setInput_State(0,delay_before=1)            #輸出OFF
        self.DUT.closePort()
        PSW.closePort()
        sJudge=self.JudgeResult(strIName)   #判斷此測試PASS或FAIL
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge
   
    #Cut off 電流準確度
    def I_CutOff_Accuracy(self,strIName='',FixUse='all'): 
        subProcName="I_CutOff_Accuracy"
        TH=recordHeader()
        sJudge=''
        if self.DUT.instrConnected==0 or  DMM.instrConnected==0:
            self.txtMsg_append(f"無法測試{strIName},待測機或DMM未連線",Qt.red)
            self.delay(0.5)
            return sJudge
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        self.DUT.openPort()        
        self.DUT.CLS()
        try:
            DMM.openPort()
        except Exception as e:
            self.txtMsg_append(f"{inspect.currentframe().f_code.co_name} Error: {str(e)}", Qt.red)
        #self.delay(0.1)        
        strM=''
        TMode=''        
        GMode=''
        TMode_l=''
        GMode_l=''
        try:
            for i in range(0,window.TRecord.model.rowCount()):             
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過      
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                #print(f"FixUse=>{FixUse}, r.value(TH.col_Cond_Fixture)=>{r.value(TH.col_Cond_Fixture)}")
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.   
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r) 
                rComment=''
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('DC')>=0:
                    TMode='DC'
                elif strM.find('AC50HZ')>=0:
                    TMode='AC50HZ'
                elif strM.find('AC60HZ')>=0:
                    TMode='AC60HZ'

                if strM.find('浮空')>=0:
                    GMode='浮空'
                else:
                    GMode='接地'
                    
                if TMode_l!=TMode:
                    if TMode=='DC':
                        DMM.setDCI('0.1') #100mA檔
                        DMM.setRefleshRate('F')
                        DMM.setDCIRESolution('MAX') #解析設最差
                        self.DUT.setEditMode('DCW',0,0.02)
                        self.DUT.setDCW_LoSet(0)
                        # self.DUT.setDCW_HiSet(r.value(TH.col_Cond_2))
                        self.DUT.setDCW_Timer(60)                    
                    else:
                        DMM.setACI('0.1') #100mA檔
                        DMM.setRefleshRate('F')
                        DMM.setACIRESolution('MAX') #解析設最差
                        self.DUT.setEditMode('ACW',0,0.02)
                        self.DUT.setACW_LoSet(0)
                        # self.DUT.setACW_HiSet(r.value(TH.col_Cond_2))  
                        self.DUT.setACW_Timer(60)                    
                        if TMode=='AC50HZ':
                            self.DUT.setACW_Freq(50)
                        elif TMode=='AC60HZ':
                            self.DUT.setACW_Freq(60)
                    TMode_l=TMode
                    self.delay(0.05)
                if GMode_l!=GMode:
                    if GMode== '浮空':
                        self.DUT.setGroundMode(0,0,0.05)
                    else: 
                        self.DUT.setGroundMode(1,0,0.05)
                    GMode_l=GMode 
                rdg_mA=''
                rdg_dut=''
                rComment=''
                rdg_D1=0            
                rdg_DF1=-999999
                        
                ae=r.value(TH.col_UpLimit)-r.value(TH.col_Expect)   #允許單邊公差
                au=r.value(TH.col_Expect)+ae*10     #暫訂DUT讀值要處範圍為10倍公差內
                al=r.value(TH.col_Expect)-ae*10
                if r.value(TH.col_Expect)>=10:
                    RampT=6
                else:
                    RampT=4
                # RampT=round(r.value(TH.col_Expect)/10+3,3)
                self.DUT.setRampTime(RampT)#設定上升時間
                
                VStart=r.value(TH.col_Cond_1)
                VStart=float(round(VStart*0.8,3))
                VStep=float(round(VStart*0.02,3))
                HStart=r.value(TH.col_Cond_2)
                HStep=float(round(HStart*0.02,3))
                VLim=float(round(r.value(TH.col_Cond_1)*1.3,3))               
                HLim=float(round(r.value(TH.col_Cond_2)*0.7,3))
                if HLim<0.001:HLim=0.001
                if TMode=='DC':
                    if VLim>6.1: VLim=6.1
                else:
                    if VLim>5: VLim=5
                V,H=self.Adj_Volt_Hiset_2Fail(TMode,VStart, VStep,HStart,HStep,VLim,HLim)   #先找到cut off電壓及Hiset設定
                self.txtMsg_append(f"取得測試電壓:{V}, HiSet:{H}",QColor('darkorchid'))
                if TMode=='DC':
                    if H<10:
                        DMM.setDCI('0.01') #10mA檔
                    else:
                        DMM.setDCI('0.1') #100mA檔
                    DMM.setRefleshRate('F')
                    DMM.setDCIRESolution('MAX') #解析設最差
                    self.DUT.setDCW_HiSet(H)
                    #self.DUT.setDCW_Volt(round(V*0.95,3))
                    self.DUT.setDCW_Volt(float(round(V,3)))
                else:
                    if H<10:
                        DMM.setACI('0.01') #10mA檔
                    else:
                        DMM.setACI('0.1') #100mA檔
                    DMM.setRefleshRate('F')
                    DMM.setACIRESolution('MAX') #解析設最差
                    self.DUT.setACW_HiSet(H)  
                    self.DUT.setACW_Volt(float(round(V,3))) 
                
                for j in range(0,3):    #測數次, DMM最大值
                    self.DUT.setRampTime(RampT)#設定上升時間
                    DMM.setCalculateState('on')
                    DMM.setCalculateFunction('MAX')
                    self.delay(0.3)
                    self.DUT.Test(1)
                    out_st=monotonic()
                    while (monotonic()-out_st)<(RampT+3):
                        self.delay(0.2)
                        MeasDut=self.DUT.getMeasure()
                        rdgD=DMM.getCalculateMax()
                        if MeasDut['Status']=="ERROR" or MeasDut['Status']=="FAIL": break
                        if self.StopTest: break                    
                    self.delay(0.3)
                    rdgD=DMM.getCalculateMax()
                    MeasU=self.DUT.getMeasure()                
                    DMM.setCalculateFunction('OFF') 
                    DMM.setCalculateState('off') 
                    if isnumber(rdgD) and isnumber(MeasU['TAR']):                        
                        rdgD=rdgD*1000      #轉mA
                        rdgU=float(MeasU['TAR'])
                        rdg_DF=rdgU-rdgD
                        self.txtMsg_append(f"{j+1}次讀值 DMM:{rdgD}, DUT:{MeasU['TAR']}")                    
                        rComment+=f"DMM {j+1}: {rdgD}, DUT:{MeasU['TAR']}" 
                        if abs(rdg_DF1)>abs(rdg_DF):
                            rdg_DF1=rdg_DF
                            rdg_D1=rdgD
                            MeasDut=MeasU
                            rdg_dut=rdgU
                    self.DUT.Test(0)
                    self.delay(0.1)
                    if abs(rdg_DF1)<abs(ae): break      #已經在規格內, 先離開
                    RampT+=5 #上升時間再加5秒
                if rdg_D1>0: rdg_mA=rdg_D1
                rComment+=f'測試電壓:{MeasDut["TVA"]}'
                r.setValue(TH.col_Comment, rComment)
                self.txtMsg_append(f"DMM讀值=> {str(rdg_mA)} mA")
                if rdg_dut!='' and rdg_mA!='' and rdg_dut<au and rdg_dut>al:
                    rdg=r.value(TH.col_Expect)+(rdg_dut-rdg_mA)
                    r.setValue(TH.col_Measure_Main,rdg)
                    r.setValue(TH.col_Measure_1,rdg_mA)
                    r.setValue(TH.col_Measure_2,rdg_dut)
                else:
                    r.setValue(TH.col_Measure_Main,'')
                    r.setValue(TH.col_Measure_1,rdg_mA)
                    r.setValue(TH.col_Measure_2,rdg_dut) 
                rComment+=",M1:DMM讀值,M2:DUT讀值"                        
                if MeasDut['Status']!="ERROR" and r.value(TH.col_Measure_Main)!='':
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                    if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                        r.setValue(TH.col_Judge, 'PASS')                    
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                                  
                else:
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                    r.setValue(TH.col_Judge, 'FAIL')
                    rComment+=f"Status: {MeasDut['Status']}"
                r.setValue(TH.col_Comment, rComment)
                self.Update_TRecord(r,True,rStartTime)            
                DErr=self.DUT.getSysErr() 
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)        
        if TMode=='DC':
            DMM.setDCIRESolution('MIN')
        else:
            DMM.setACIRESolution('MIN')
        self.delay(0.1)
        self.DUT.setRampTime(0.1)            
        self.DUT.Test(0)
        self.DUT.closePort()        
        DMM.closePort()
        sJudge=self.JudgeResult(strIName) 
        window.TRecord.model.TestRow=-1     
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e
        QtWidgets.QApplication.processEvents()         
        return sJudge
    
    #調電壓, Hiset到cut off.
    def Adj_Volt_Hiset_2Fail(self, TMode, VStart, VStep, HisetStart, HisetStep, VLimit=5.1, HisetLimit=0.001):
        subProcName='adjVolt2Fail'
        v=VStart
        last_v=v
        hiset=HisetStart
        try:          
            self.DUT.CLS()
            self.delay(0.1)
            self.DUT.Test(0)
            RampT=0.5
            self.DUT.setRampTime(RampT)#設定上升時間
            if str(TMode).find('DC')>=0:
                self.DUT.setDCW_HiSet(HisetStart)
            else:
                self.DUT.setACW_HiSet(HisetStart) 
            i=0
            #調電壓
            ErrResetV=False
            while (i<10 and not ErrResetV):                 
                self.DUT.Test(0)
                self.delay(0.1)            
                if str(TMode).find('DC')>=0:                    
                    self.DUT.setDCW_Volt(v)                                       
                else:
                    self.DUT.setACW_Volt(v)  
                last_v=v   #最後電壓設定值 
                self.txtMsg_append(f"{i}次, 設定電壓:{last_v}")
                self.delay(0.2)
                self.DUT.Test(1) 
                out_st=monotonic()
                k=0
                while (monotonic()-out_st)<(RampT+1) and k<5:
                    self.delay(0.5)
                    MeasDut=self.DUT.getMeasure()  
                    DErr=self.DUT.getSysErr()   #取得系統錯誤訊息.                                                            
                    if MeasDut['Status']=="ERROR" or DErr!="":
                        self.txtMsg_append(f"k:{k},發現錯誤({DErr})")
                        self.DUT.Test(0)
                        if DErr!="":  
                            if str(DErr).lower().find('dc over')>=0:    #超過DC功率限制, 調降電壓
                                w=[100,50]
                                for sw in w:                                    
                                    v=float(round(sw/HisetStart,3))-0.01
                                    if str(TMode).find('DC')>=0:
                                        self.DUT.setDCW_Volt(v)
                                    else:
                                        self.DUT.setACW_Volt(v)
                                    self.delay(0.1)
                                    DErr=self.DUT.getSysErr()   #取得系統錯誤訊息.
                                    if DErr=="": break
                            else:
                                try:
                                    v=VPre
                                except:
                                    v-=VStep   
                            v=float(round(v,3))                         
                            if str(TMode).find('DC')>=0:
                                self.DUT.setDCW_Volt(v)
                            else:
                                self.DUT.setACW_Volt(v)
                            last_v=v   #最後電壓設定值
                            self.txtMsg_append(f"重設電壓:{v}")
                            ErrResetV=True                            
                            self.DUT.Test(1)                            
                            out_st=monotonic()
                            k+=1                    
                    if MeasDut['Status'] in ["ERROR","FAIL","ARC"]: break
                    if self.StopTest: break  
                self.txtMsg_append(f"DUT狀態:{MeasDut['Status']}, 電壓:{str(MeasDut['TVA'])}, 電流:{str(MeasDut['TAR'])}")                                      
                if self.StopTest: break
                if v>=VLimit:
                    self.txtMsg_append(f"調整電壓值達上限, 電壓:{v}, 上限:{VLimit}")
                    break
                if MeasDut['Status']=="ERROR" or MeasDut['Status']=="FAIL":break 
                VPre=v
                if i==0:
                    R1=MeasDut['TVA']/MeasDut['TAR']
                    v=hiset*R1*1.01 #調高計算值多1%
                    if v>=VLimit: v=VLimit  
                    if v<VStart:V=VStart
                    v=float(round(v,3))
                    self.txtMsg_append(f"計算電阻值:{float(round(R1*1000,3))}Kohm, 電壓將改為:{v}KV")                   
                elif not ErrResetV:
                    v+=VStep
                    v=float(round(v,3))
                    self.txtMsg_append(f"未達Cut off, 將調高電壓一階為:{v}KV") 
                if v>VLimit: v=VLimit
                i+=1
            
            #調低Hiset
            i=0
            while str(MeasDut['Status']).upper().strip()!="FAIL" and i<10: 
                self.DUT.CLS()
                self.delay(0.1)
                MeasDut=self.DUT.getMeasure()
                self.delay(0.1)               
                self.DUT.Test(0)
                self.delay(0.1)
                if i==0:
                    R1=MeasDut['TVA']/MeasDut['TAR']                    
                    hiset=round(v/R1*0.98,3)
                    self.txtMsg_append(f"計算電阻值:{round(R1*1000,3)}Kohm,, hiset將改為:{hiset}mA")
                else:
                    hiset-=HisetStep 
                if hiset<HisetLimit: hiset=HisetLimit         
                if str(TMode).find('DC')>=0:
                    self.DUT.setDCW_Volt(v)                                  
                    self.DUT.setDCW_HiSet(hiset)
                else:
                    self.DUT.setACW_Volt(v)
                    self.DUT.setACW_HiSet(hiset)
                self.txtMsg_append(f"重設定電壓:{v}, HiSet:{hiset}")
                self.delay(0.1)
                self.DUT.Test(1)                            
                out_st=monotonic()
                while (monotonic()-out_st)<(RampT+1):
                    self.delay(0.2)
                    MeasDut=self.DUT.getMeasure()                    
                    if MeasDut['Status']=="ERROR" or MeasDut['Status']=="FAIL": break
                    if self.StopTest: break
                if self.StopTest: break 
                if hiset <=HisetLimit: 
                    self.txtMsg_append(f"重設定HiSet達到下限, HiSet:{hiset}mA")
                    break 
                i+=1  
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.Test(0)                
        return last_v,hiset
    
    #IR電壓準確度
    def PQC_IR_V_Accuracy(self,strIName='',FixUse='all'):
        subProcName="PQC_IR_V_Accuracy"   
        TH=recordHeader()
        sJudge=''        
        if self.DUT.instrConnected==0 or  not (HVM.instrConnected==1 or self.chkHVMManual.isChecked()):
            self.txtMsg_append(f"無法測試{strIName},待測機或高壓表未連線",Qt.red)
            self.delay(0.5)
            return sJudge
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge        
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()
            HVM.openPort()
            self.DUT.CLS()
            initSet=True
            for i in range(0,window.TRecord.model.rowCount()): 
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過           
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        rHiSet=cond_DUT['HISET']
                    except:
                        rHiSet=''
                except:
                    rHiSet=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                self.DUT.CLS()      #清除DUT錯誤訊息
                if initSet:                     #第一次設定後, 不再設定           
                    self.DUT.setEditMode('IR',0,0.02)      #設定IR MODE
                    self.DUT.setIR_Timer(60)
                    # self.DUT.setIR_HiSet(9999)
                    if fnmatch(self.cboScript.currentText(),'*120*') or fnmatch(self.cboScript.currentText(),'*150*'):
                        self.DUT.setIR_LoSet(0.1)
                    else:
                        self.DUT.setIR_LoSet(1)
                    if fnmatch(str(self.DUT.getIR_LoSet()),'*G*'): #說明書單位為M ohm, 但發現有實機下'1', 會變成1G ohm
                        self.DUT.setIR_LoSet(0.001)
                    initSet=False
                if rHiSet!='':
                    self.DUT.setIR_HiSet(rHiSet)
                else:
                    self.DUT.setIR_HiSet('NULL')
                self.DUT.setIR_Volt(float(r.value(TH.col_Cond_1))/1000)
                self.delay(0.2,stm)
                self.DUT.Test(1)                #輸出                        
                self.delay(1)                   #等待
                MeasDut=self.DUT.getMeasure()
                if MeasDut['Status']!="ERROR" and MeasDut['Status']!="FAIL"  and MeasDut['Status']!="STOP":
                    self.delay(1.5)               #没有發生ERROR
                if self.chkHVMManual.isChecked()==False:
                    rComment=''
                    r.setValue(TH.col_Comment, rComment)
                    rdg_hvm=HVM.getDCV()
                else:
                    if not self.StopTest:
                        rComment='手填值 '
                        r.setValue(TH.col_Comment, rComment)
                        rdg_hvm=self.HVM_Manual()
                r.setValue(TH.col_Measure_Main,rdg_hvm)
                self.txtMsg_append(f"高壓表讀值=> {str(rdg_hvm)} V")
                MeasDut=self.DUT.getMeasure()            
                if MeasDut['Status']!="ERROR" and rdg_hvm!='':
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                    if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                        r.setValue(TH.col_Judge, 'PASS')                    
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                                  
                else:
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                    r.setValue(TH.col_Judge, 'FAIL')
                    rComment+=f"{MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                self.Update_TRecord(r,True,rStartTime)
                self.DUT.Test(0)        #輸出OFF
                self.delay(0.1)
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.Test(0)        #輸出OFF
        self.DUT.closePort()
        HVM.closePort()
        sJudge=self.JudgeResult(strIName)     
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e 
        return sJudge
    
    #IR絕緣阻抗量測準確度
    def IR_Meas_Accuracy(self,strIName='',FixUse='all'):
        subProcName="IR_Meas_Accuracy"
        TH=recordHeader()
        sJudge=''        
        if self.DUT.instrConnected==0:
            self.txtMsg_append(f"無法測試{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge  
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge      
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()        
            self.DUT.CLS()
            initSet=True
            for i in range(0,window.TRecord.model.rowCount()): 
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過           
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        rHiSet=cond_DUT['HISET']
                    except:
                        rHiSet=''
                except:
                    rHiSet=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                self.DUT.CLS()      #清除DUT錯誤訊息
                if initSet:                     #第一次設定後, 不再設定           
                    self.DUT.setEditMode('IR',0,0.02)      #設定IR MODE
                    self.DUT.setIR_Timer(60)
                    if fnmatch(self.cboScript.currentText(),'*120*') or fnmatch(self.cboScript.currentText(),'*150*'):
                        self.DUT.setIR_LoSet(0.1)
                    else:
                        self.DUT.setIR_LoSet(1)
                    if fnmatch(str(self.DUT.getIR_LoSet()),'*G*'): #說明書單位為M ohm, 但發現有實機下'1', 會變成1G ohm
                        self.DUT.setIR_LoSet(0.001)
                    initSet=False
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t            
                 
                if strM.find('接地')>=0:
                    self.DUT.setGroundMode(1,0,0.02)
                else:
                    self.DUT.setGroundMode(0,0,0.02)
                
                if rHiSet!='':
                    self.DUT.setIR_HiSet(rHiSet)
                else:
                    self.DUT.setIR_HiSet('NULL')
                self.DUT.setIR_Volt(float(r.value(TH.col_Cond_1))/1000) #轉成KV
                self.delay(0.2,stm)
                self.DUT.Test(1)                #輸出                        
                self.delay(1)                   #等待
                MeasDut=self.DUT.getMeasure()
                if r.value(TH.col_Expect)>1000:
                    dlt=3
                else:
                    dlt=1
                if MeasDut['Status']!="ERROR" and MeasDut['Status']!="FAIL":
                    self.delay(dlt)               #没有發生ERROR
                MeasDut=self.DUT.getMeasure()
                if isnumber(MeasDut['TAR']):                
                    rdg_dut=float(MeasDut['TAR'])   
                else:
                    rdg_dut=''

                rdg_std, rdg_std_unit=self.getFixCalValue(r.value(TH.col_Cond_Fixture))

                ae=r.value(TH.col_UpLimit)-r.value(TH.col_Expect)   #允許單邊公差
                au=r.value(TH.col_Expect)+ae*5     #暫訂DUT讀值要處範圍為5倍公差內
                al=r.value(TH.col_Expect)-ae*5                
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                r.setValue(TH.col_Measure_2,rdg_dut)
                r.setValue(TH.col_Measure_1,rdg_std)                 
                if rdg_std=='':
                    sFPathNoFound=f"找不到治具路徑({r.value(TH.col_Cond_Fixture)})的校正值"
                    rComment+=f"{sFPathNoFound}, "
                rComment+="M1:治具阻值,M2:DUT讀值"           
                try:
                    if r.value(TH.col_Measure_2)>al and r.value(TH.col_Measure_2)<au and rdg_std!='':
                        d=r.value(TH.col_Expect)+(rdg_dut-rdg_std)
                    else:
                        d=''                
                except:
                    d=''
                r.setValue(TH.col_Measure_Main,d)
                if MeasDut['Status']!="ERROR" and r.value(TH.col_Measure_Main)!='':
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                    if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                        r.setValue(TH.col_Judge, 'PASS')                    
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                                  
                else:
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                    r.setValue(TH.col_Judge, 'FAIL')
                    rComment+=f"{MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                self.Update_TRecord(r,True,rStartTime)
                self.DUT.Test(0)        #輸出OFF
                self.delay(0.1)
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort()
        sJudge=self.JudgeResult(strIName)   
        window.TRecord.model.TestRow=-1  
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e
        return sJudge

    #GB電流設定準確度
    def GB_I_Set_Accuracy(self,strIName='',FixUse='all'):
        subProcName="GB_I_Set_Accuracy"
        TH=recordHeader()
        sJudge=''        
        if self.DUT.instrConnected==0 or PCS.instrConnected==0:
            self.txtMsg_append(f"無法測試{strIName},待測機或PCS未連線",Qt.red)
            self.delay(0.5)
            return sJudge  
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge      
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            PCS.openPort()
            PCS.CLS()
            PCS.setACI(30)          #設PCS AC30A
            self.DUT.openPort()
            self.DUT.CLS()
            initSet=True
            TFreq_1=''
            for i in range(0,window.TRecord.model.rowCount()): 
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過           
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        rHiSet=cond_DUT['HISET']
                    except:
                        rHiSet=''
                except:
                    rHiSet=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                self.DUT.CLS()      #清除DUT錯誤訊息
                if initSet:                     #第一次設定後, 不再設定           
                    self.DUT.setEditMode('GB',0,0.02)      #設定IR MODE
                    self.DUT.setGB_Timer(60)
                    self.DUT.setGB_LoSet(0)
                    GBHiset=self.default_GB_HiSet
                    # self.DUT.setGB_HiSet(GBHiset)
                    initSet=False
                if rHiSet!='':
                    self.DUT.setGB_HiSet(rHiSet) 
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t            
                # strM=i2t
                if strM.find('50HZ')>=0:
                    TFreq=50            
                elif strM.find('60HZ')>=0:
                    TFreq=60 
                if strM.find('接地')>=0:
                    self.DUT.setGroundMode(1,0,0.02)
                else:
                    self.DUT.setGroundMode(0,0,0.02)
                if TFreq_1!=TFreq:
                    self.DUT.setGB_Freq(TFreq)      #設GB頻率
                    TFreq_1!=TFreq
                GBCur=float(r.value(TH.col_Cond_1))
                if GBCur*GBHiset/1000>self.GBVLimit:        #超出允許電壓, 調低阻值
                    GBHiset=int(self.GBVLimit/GBCur*1000)
                self.DUT.setGB_HiSet(GBHiset)
                self.DUT.setGB_Curr(GBCur) #設GB電流
                self.delay(0.3,stm)
                self.DUT.Test(1)                #輸出                        
                self.delay(1)                   #等待
                MeasDut=self.DUT.getMeasure()
                if MeasDut['Status']!="ERROR" and MeasDut['Status']!="FAIL":
                    self.delay(1.5)               #没有發生ERROR
                MeasDut=self.DUT.getMeasure()
                if isnumber(MeasDut['TVA']):                
                    rdg_dut=float(MeasDut['TVA'])   
                else:
                    rdg_dut=''

                rdg_std=PCS.getReading_ACI()
                ae=r.value(TH.col_UpLimit)-r.value(TH.col_Expect)        #允許單邊公差
                au=r.value(TH.col_Expect)+ae*5              #暫訂DUT讀值要處範圍為5倍公差內
                al=r.value(TH.col_Expect)-ae*5                
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                r.setValue(TH.col_Measure_2,rdg_dut)
                r.setValue(TH.col_Measure_1,rdg_std)
                rComment+="M1:PCS讀值,M2:DUT讀值"            
                try:
                    if r.value(TH.col_Measure_2)>al and r.value(TH.col_Measure_2)<au:
                        d=r.value(TH.col_Expect)+(rdg_dut-rdg_std)
                    else:
                        d=''                
                except:
                    d=''
                r.setValue(TH.col_Measure_Main,d)
                if MeasDut['Status']!="ERROR" and r.value(TH.col_Measure_Main)!='':
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                    if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                        r.setValue(TH.col_Judge, 'PASS')                    
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                                  
                else:
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                    r.setValue(TH.col_Judge, 'FAIL')
                    rComment+=f"{MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                self.Update_TRecord(r,True,rStartTime)
                self.DUT.Test(0)        #輸出OFF
                self.delay(0.1)
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort()
        PCS.closePort()
        sJudge=self.JudgeResult(strIName)   
        window.TRecord.model.TestRow=-1 
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e
        return sJudge
    
    #GB電阻量測準確度
    def GB_R_Meas_Accuracy(self,strIName='',FixUse='all'):
        subProcName="GB_R_Meas_Accuracy"
        # self.txtMsg.setText('GB電阻量測準確度')
        TH=recordHeader()
        sJudge=''        
        if self.DUT.instrConnected==0:
            self.txtMsg_append(f"無法測試{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge  
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge      
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()
            self.DUT.CLS()
            initSet=True
            TFreq_1=''
            for i in range(0,window.TRecord.model.rowCount()): 
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過           
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        rHiSet=cond_DUT['HISET']
                    except:
                        rHiSet=''
                except:
                    rHiSet=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                self.delay(0.5)     #等開闗切上
                stm=monotonic()     #設等待的時間起點
                self.DUT.CLS()      #清除DUT錯誤訊息
                if initSet:                     #第一次設定後, 不再設定           
                    self.DUT.setEditMode('GB',0,0.02)      #設定IR MODE
                    self.DUT.setGB_Timer(60)
                    self.DUT.setGB_LoSet(0)
                    GBHiset=self.default_GB_HiSet
                    # self.txtMsg_append(f"設定GB HiSet: {GBHiset}")
                    # self.DUT.setGB_HiSet(GBHiset)
                    # self.delay(0.1)
                    # self.Show_DUT_SysErr()
                    initSet=False
                # if rHiSet!='':
                if isnumber(rHiSet):
                    if float(rHiSet)>(r.value(TH.col_Expect)+30):
                        GBHiset=float(rHiSet)   
                        # self.DUT.setGB_HiSet(rHiSet) 
                    else:
                        GBHiset=r.value(TH.col_Expect)+30 
                        # self.DUT.setGB_HiSet(r.value(TH.col_Expect)+30)                   
                else:
                    GBHiset=r.value(TH.col_Expect)+30 
                    # self.DUT.setGB_HiSet(r.value(TH.col_Expect)+30)
                self.txtMsg_append(f"設定GB HiSet: {GBHiset}")
                self.DUT.setGB_HiSet(GBHiset)
                self.delay(0.1)
                self.Show_DUT_SysErr()
                
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")  
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t          
                # strM=i2t
                if strM.find('50HZ')>=0:                    
                    TFreq=50            
                elif strM.find('60HZ')>=0:
                    TFreq=60 

                if strM.find('接地')>=0:
                    self.txtMsg_append(f"設定GB Ground Mode: 接地")
                    self.DUT.setGroundMode(1,0,0.02)
                    self.delay(0.1)
                    self.Show_DUT_SysErr()
                else:
                    self.txtMsg_append(f"設定GB Ground Mode: 浮空")
                    self.DUT.setGroundMode(0,0,0.02)  
                    self.delay(0.1)
                    self.Show_DUT_SysErr()
                
                if TFreq_1!=TFreq:
                    self.txtMsg_append(f"設定GB Freq:{TFreq}")
                    self.DUT.setGB_Freq(TFreq)      #設GB頻率
                    self.delay(0.1)
                    self.Show_DUT_SysErr()
                    TFreq_1!=TFreq
                GBCur=float(r.value(TH.col_Cond_1))
                if GBCur*GBHiset/1000>self.GBVLimit:        #超出允許電壓, 調低阻值
                    GBHiset=int(self.GBVLimit/GBCur*1000)
                    self.txtMsg_append(f"設定GB Hiset:{GBHiset}")
                    self.DUT.setGB_HiSet(GBHiset)
                    self.delay(0.1)
                    self.Show_DUT_SysErr()

                self.txtMsg_append(f"設定GB電流:{GBCur}")
                self.DUT.setGB_Curr(GBCur) #設GB電流                
                self.delay(0.3,stm)
                self.Show_DUT_SysErr()

                self.DUT.Test(1)                #輸出 
                self.txtMsg_append(f"{QDateTime.currentDateTime().toString(self.MyTimeFormart)} GB_R測試輸出...") 
                                      
                # self.delay(1.5)                   #等待輸出穩定
                self.delay(2)                   #等待輸出
                MeasDut=self.DUT.getMeasure()
                self.txtMsg_append(f"{QDateTime.currentDateTime().toString(self.MyTimeFormart)} GB_R讀值:{MeasDut['TAR']}")
                rd=-1       #隨便設定不應有的值為初始值
                rd1=-2      #隨便設定不應有的值為初始值
                if isnumber(MeasDut['TAR']):                
                    rd=float(MeasDut['TAR'])  
                if MeasDut['Status']!="ERROR":
                    self.txtMsg_append("DUT Error Message:",Qt.red)
                    self.txtMsg_append(self.DUT.getErrMessage())
                    self.txtMsg_append(self.DUT.getSysErr())
                    QApplication.processEvents()
                # if MeasDut['Status']!="ERROR" and MeasDut['Status']!="FAIL" and isnumber(MeasDut['TAR']):   #没有發生ERROR, 限定時間內, 讀值讀到穩定值
                if MeasDut['Status']!="FAIL" :
                    for ir in range(0,5):
                        self.delay(1)
                        MeasDut=self.DUT.getMeasure()
                        self.txtMsg_append(f"{QDateTime.currentDateTime().toString(self.MyTimeFormart)} GB_R讀值:{MeasDut['TAR']}")
                        if isnumber(MeasDut['TAR']):                
                            rd1=float(MeasDut['TAR'])
                            if rd1==rd:
                                break
                            else:
                                rd=rd1
                        else:
                            break   #得到的值, 不能轉數值, 可能發生錯誤, 跳離

                MeasDut=self.DUT.getMeasure()
                self.txtMsg_append(f"{QDateTime.currentDateTime().toString(self.MyTimeFormart)} GB_R讀值:{MeasDut['TAR']}")
                if isnumber(MeasDut['TAR']):                
                    rdg_dut=float(MeasDut['TAR'])   
                else:
                    rdg_dut=''
                rdg_std, rdg_std_unit=self.getFixCalValue(r.value(TH.col_Cond_Fixture),r.value(TH.col_Cond_1))
                ae=r.value(TH.col_UpLimit)-r.value(TH.col_Expect) +25   #允許單邊公差,+25m ohm給Relay接點抗阻用
                au=r.value(TH.col_Expect)+ae*5     #暫訂DUT讀值要處範圍為5倍公差內
                al=r.value(TH.col_Expect)-ae*5
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                r.setValue(TH.col_Measure_2,rdg_dut)
                r.setValue(TH.col_Measure_1,rdg_std)
                if rdg_std=='':
                    sFPathNoFound=f"找不到治具路徑({r.value(TH.col_Cond_Fixture)})的校正值"
                    rComment+=f"{sFPathNoFound}, "
                rComment+="M1:治具阻值,M2:DUT讀值"            
                try:
                    if r.value(TH.col_Measure_2)>al and r.value(TH.col_Measure_2)<au and rdg_std!='':
                        d=r.value(TH.col_Expect)+(rdg_dut-rdg_std)
                    else:
                        d=''                
                except:
                    d=''
                r.setValue(TH.col_Measure_Main,d)
                if MeasDut['Status']!="ERROR" and r.value(TH.col_Measure_Main)!='':
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                    if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                        r.setValue(TH.col_Judge, 'PASS')                    
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                                  
                else:
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                    r.setValue(TH.col_Judge, 'FAIL')
                    rComment+=f"{MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                self.Update_TRecord(r,True,rStartTime)
                self.DUT.Test(0)        #輸出OFF
                self.delay(0.1)
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort()
        sJudge=self.JudgeResult(strIName) 
        window.TRecord.model.TestRow=-1  
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e  
        return sJudge
    
    #查詢是否有SysErr, 若有顯示出來
    def Show_DUT_SysErr(self):
        strErr=self.DUT.getSysErr()
        if strErr!="":
            self.txtMsg_append(f"DUT Sys Err:{strErr}",Qt.red)

    #空載漏電流確認
    def I_Leakage(self,FixUse='all'):
        subProcName="I_Leakage"
        strIName=str("空載漏電流確認")        
        TH=recordHeader()        
        sJudge=''                
        if self.DUT.instrConnected==0 :
            self.txtMsg_append(f"無法進行{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge  
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge      
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()        
            self.DUT.CLS()
            for i in range(0,window.TRecord.model.rowCount()):
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過                      
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue 
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r) 
                rComment=''
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        rHiSet=cond_DUT['HISET']
                    except:
                        rHiSet=''
                except:
                    rHiSet=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                strM=i2t+i3t
                if strM.find('AC')>=0: 
                    self.DUT.setEditMode('ACW',0,0.02)
                    if rHiSet=='':
                        self.DUT.setACW_HiSet(0.1)  
                    else:
                        self.DUT.setACW_HiSet(rHiSet)
                    self.DUT.setACW_Timer(60) 
                    if strM.find('50')>=0 or str(r.value(TH.col_Cond_2)).find('50')>=0:                 
                        self.DUT.setACW_Freq_Volt('50',r.value(TH.col_Cond_1))
                    else:
                        self.DUT.setACW_Freq_Volt('60',r.value(TH.col_Cond_1))
                elif strM.find('DC')>=0: 
                    self.DUT.setEditMode('DCW',0,0.02)
                    if rHiSet=='':
                        self.DUT.setDCW_HiSet(0.1)  
                    else:
                        self.DUT.setDCW_HiSet(rHiSet)
                    self.DUT.setDCW_Timer(60)
                    self.DUT.setDCW_Volt(r.value(TH.col_Cond_1)) 
                self.delay(0.05)
                if strM.find('浮空')>=0:
                    self.DUT.setGroundMode(0)
                else: 
                    self.DUT.setGroundMode(1)
                
                self.delay(0.2,stm)
                self.DUT.Test(1)
                        
                self.delay(1)                   #等待            
                MeasDut=self.DUT.getMeasure()
                if MeasDut['Status']!="ERROR":               
                    self.delay(3)
                    MeasDut=self.DUT.getMeasure()
                    mV=MeasDut['TAR']
                    r.setValue(TH.col_Measure_Main,mV)
                    if isnumber(mV):                        
                        if r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                            r.setValue(TH.col_Judge, 'PASS')                    
                        else:
                            r.setValue(TH.col_Judge, 'FAIL')
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                    
                    self.txtMsg_append(f"測試狀態:{MeasDut['Status']}, 漏電流:{MeasDut['TAR']}mA")
                    self.txtMsg_append(str(MeasDut))
                    self.txtMsg_append(f"DUT最後回覆訊息:{self.DUT.LastAnswer}")
                else:
                    self.txtMsg_append(f"測試狀態:{MeasDut['Status']}, 漏電流:{MeasDut['TAR']}mA",Qt.red)
                    self.txtMsg_append(str(MeasDut))
                    self.txtMsg_append(f"DUT最後回覆訊息:{self.DUT.LastAnswer}")
                    r.setValue(TH.col_Judge,'FAIL')
                    rComment+=f"{MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                    
                self.DUT.Test(0)    #輸出OFF
                self.delay(0.1)
                DErr=self.DUT.getSysErr()
                # DErr=str(DErr).strip().replace("'","").replace("[","").replace("]","")
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
                self.Update_TRecord(r,True,rStartTime)
                sIndex = self.tbView.model().index(i, 0)
                self.tbView.scrollTo(sIndex)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort() 
        sJudge=self.JudgeResult(strIName)
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge
    
    #CONT量測準確度
    def CONT_Meas_Accuracy(self,strIName='',FixUse='all'):
        subProcName="CONT_Meas_Accuracy"
        TH=recordHeader()
        sJudge=''        
        if self.DUT.instrConnected==0:
            self.txtMsg_append(f"無法測試{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge  
        if not(self.chkFixManual.isChecked()) and DIO.instrConnected==0:                #非手動換線又治具未連線
            self.txtMsg_append(f"無法測試{strIName},未勾選手動接線治具且治具未連線",Qt.red)
            self.delay(0.5)
            return sJudge      
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()        
            self.DUT.CLS()
            initSet=True
            for i in range(0,window.TRecord.model.rowCount()): 
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過           
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue     #如果治具使用不是'all'且不是'對應的治具', 執行下一回圈.
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r)
                rComment=''
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        rHiSet=cond_DUT['HISET']
                    except:
                        rHiSet=''
                except:
                    rHiSet=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                self.DUT.CLS()      #清除DUT錯誤訊息
                if initSet:                     #第一次設定後, 不再設定           
                    self.DUT.setEditMode('CONT',0,0.02)      #設定IR MODE
                    self.DUT.setCONT_Timer(60)
                    self.DUT.setCONT_LoSet(0.01)                    
                    initSet=False
                if rHiSet!='':
                    self.DUT.setCONT_HiSet(rHiSet)
                else:
                    self.DUT.setCONT_HiSet(80)                
                self.delay(0.2,stm)
                self.DUT.Test(1)                #輸出                        
                self.delay(1)                   #等待
                MeasDut=self.DUT.getMeasure()
                if r.value(TH.col_Expect)>1000:
                    dlt=3
                else:
                    dlt=1
                if MeasDut['Status']!="ERROR" and MeasDut['Status']!="FAIL":
                    self.delay(dlt)               #没有發生ERROR
                MeasDut=self.DUT.getMeasure()
                if isnumber(MeasDut['TAR']):                
                    rdg_dut=float(MeasDut['TAR'])   
                else:
                    rdg_dut=''
                rdg_std, rdg_std_unit=self.getFixCalValue(r.value(TH.col_Cond_Fixture))
                ae=r.value(TH.col_UpLimit)-r.value(TH.col_Expect)   #允許單邊公差
                au=r.value(TH.col_Expect)+ae*5     #暫訂DUT讀值要處範圍為5倍公差內
                al=r.value(TH.col_Expect)-ae*5                
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                r.setValue(TH.col_Measure_2,rdg_dut)
                r.setValue(TH.col_Measure_1,rdg_std)                 
                if rdg_std=='':
                    sFPathNoFound=f"找不到治具路徑({r.value(TH.col_Cond_Fixture)})的校正值"
                    rComment+=f"{sFPathNoFound}, "
                rComment+="M1:治具阻值,M2:DUT讀值"           
                try:
                    if r.value(TH.col_Measure_2)>al and r.value(TH.col_Measure_2)<au and rdg_std!='':
                        d=r.value(TH.col_Expect)+(rdg_dut-rdg_std)
                    else:
                        d=''                
                except:
                    d=''
                r.setValue(TH.col_Measure_Main,d)
                if MeasDut['Status']!="ERROR" and r.value(TH.col_Measure_Main)!='':
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}")
                    if r.value(TH.col_Measure_Main) > r.value(TH.col_LowLimit) and r.value(TH.col_Measure_Main) < r.value(TH.col_UpLimit):
                        r.setValue(TH.col_Judge, 'PASS')                    
                    else:
                        r.setValue(TH.col_Judge, 'FAIL')                                  
                else:
                    self.txtMsg_append(f"DUT Measure=> {MeasDut['Function']}, {MeasDut['Status']}, {str(MeasDut['TVA'])}, {str(MeasDut['TAR'])}, {str(MeasDut['TT'])}",Qt.red)
                    r.setValue(TH.col_Judge, 'FAIL')
                    rComment+=f"{MeasDut['Status']}"
                    r.setValue(TH.col_Comment, rComment)
                self.Update_TRecord(r,True,rStartTime)
                self.DUT.Test(0)        #輸出OFF
                self.delay(0.1)
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort()
        sJudge=self.JudgeResult(strIName)   
        window.TRecord.model.TestRow=-1  
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e
        return sJudge


    #背板輸出確認
    def Manual_Check_Func(self,strIName='',FixUse='all'):
        subProcName="Manual_Check_Func"
        # self.txtMsg.setText('背板輸出確認')
        # strIName=str("背板輸出確認")        
        TH=recordHeader()        
        sJudge='' 
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            for i in range(0,window.TRecord.model.rowCount()):
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過                      
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue 
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r) 
                rComment='手動功能確認'
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                QtWidgets.QApplication.processEvents()
                sPt=f"{strIName}\n"
                if r.value(TH.col_ItemName_2t) !='':
                    sPt+=f"{r.value(TH.col_ItemName_2t)}"
                if r.value(TH.col_ItemName_3t)!='':
                    sPt+=f", {r.value(TH.col_ItemName_3t)}"
                if r.value(TH.col_Cond_1)!='':
                    sPt+=f", {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)}"
                sPt+=f"\n請確認是否OK, OK接Yes, NG按No."
                reply = QMessageBox.information(self, '手動功能確認', sPt, 
                                QMessageBox.No | QMessageBox.Yes, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    r.setValue(TH.col_Judge, 'PASS')
                else:
                    r.setValue(TH.col_Judge, 'FAIL')
                self.Update_TRecord(r,True,rStartTime)
                sIndex = self.tbView.model().index(i, 0)
                self.tbView.scrollTo(sIndex) 
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        sJudge=self.JudgeResult(strIName)
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e
        return sJudge

    #背板輸出確認
    def Rear_Output_Func(self,strIName='',FixUse='all'):
        # self.txtMsg.setText('背板輸出確認')
        strIName=str("背板輸出確認")        
        TH=recordHeader()        
        sJudge='' 
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()                
        sJudge=self.JudgeResult(strIName)
        return sJudge

    #掃描功能確認
    def Scan_Func(self,FixUse='all'):
        self.txtMsg.setText('掃描功能確認')

    #ARC功能確認
    def ARC_Func(self,FixUse='all'):
        self.txtMsg.setText('ARC功能確認')

    #REMOTE功能確認
    def Remote_Func(self,FixUse='all'):
        self.txtMsg.setText('REMOTE功能確認')

    #Signal IO功能確認
    def Signal_IO_Func(self,FixUse='all'):
        self.txtMsg.setText('Signal IO功能確認')

    #GPIB功能確認
    def GPIB_Func(self,FixUse='all'):
        self.txtMsg.setText('GPIB功能確認')    

    #寫入USB識別碼
    def Write_USB_ID_Check(self,strIName,FixUse='all'):
        subProcName="Write_USB_ID_Check"
        TH=recordHeader()        
        sJudge=''                
        if self.DUT.instrConnected==0 :
            self.txtMsg_append(f"無法進行{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()        
            self.DUT.CLS()
            self.delay(0.1)
            for i in range(0,window.TRecord.model.rowCount()):
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過                      
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue 
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r) 
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                try:
                    cond_DUT=eval(r.value(TH.col_Cond_DUT))
                    try:
                        USBID_Model=f"USBID_{self.ltxtDUTModel.text().upper().strip()}"
                        if (USBID_Model) in cond_DUT:
                            usbid=cond_DUT[USBID_Model]
                        else:
                            usbid=cond_DUT['USBID']
                    except:
                        usbid=''
                except:
                    usbid=''
                d=self.DUT.getDevInfo()
                
                PNPID=d['PNPDeviceID']
                PNPID_C=str(PNPID).upper().replace(" ", "").replace("USB","").replace("\\","").replace("_","").replace("&","")      #把USB字串不需的字去除.
                Need_USBID_C=usbid.upper().replace(" ", "").replace("USB","").replace("\\","").replace("_","").replace("&","")      #把USB字串不需的字去除.
                # self.txtMsg_append(f"讀回DUT VID/PID: {PNPID}")
                rComment=f"讀回VID及PID:{PNPID}"
                
                # r.setValue(TH.col_Comment, rComment)
                # if fnmatch(PNPID,f"*{usbid}*"):
                if fnmatch(PNPID_C,f"*{Need_USBID_C}*"):
                    r.setValue(TH.col_Judge, 'PASS')
                    self.txtMsg_append(rComment,Qt.blue)
                else:
                    r.setValue(TH.col_Judge, 'FAIL')
                    self.txtMsg_append(f"{rComment}, \n期望讀回:{usbid}",Qt.red)
                    self.txtMsg_append(f"以簡化期望讀回值:{Need_USBID_C}, \比對簡化讀回值:{PNPID_C}, 結果NG!",Qt.red)
                r.setValue(TH.col_Comment, rComment)
                
                DErr=self.DUT.getSysErr()
                # DErr=str(DErr).strip().replace("'","").replace("[","").replace("]","")
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
                self.Update_TRecord(r,True,rStartTime)
                sIndex = self.tbView.model().index(i, 0)
                self.tbView.scrollTo(sIndex)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort() 
        sJudge=self.JudgeResult(strIName)
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e
        return sJudge

    #寫入機器序號
    def Write_DUT_SN(self,FixUse='all'):
        subProcName="Write_DUT_SN"
        strIName=str("寫入機器序號")        
        TH=recordHeader()        
        sJudge=''                
        if self.DUT.instrConnected==0 :
            self.txtMsg_append(f"無法進行{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()        
            self.DUT.CLS()
            self.delay(0.1)
            for i in range(0,window.TRecord.model.rowCount()):
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過                      
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue 
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r) 
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                
                SN=self.ltxtDUTSn.text()                
                rSN=self.DUT.getSN()    #先讀取機器SN
                rComment=f"設定序號:{SN}, 讀回序號:{rSN}"
                
                # r.setValue(TH.col_Comment, rComment)
                if rSN==SN:             #若序號相符,PASS
                    r.setValue(TH.col_Judge, 'PASS')
                    self.txtMsg_append(rComment,Qt.blue)
                else:               #若序號不相符, 燒序號, 再讀序號比對     
                    self.DUT.CLS()
                    self.delay(0.2)
                    self.DUT.RST()
                    self.txtMsg_append("DUT Reset...",Qt.blue)
                    self.delay(1)
                    self.txtMsg_append("DUT 燒序號...",Qt.blue)
                    self.DUT.writeSN(SN)                  #等待 
                    self.delay(0.3)
                    rSN=self.DUT.getSN()
                    rComment=f"燒序號, 設定序號:{SN}, 讀回序號:{rSN}"
                    if rSN==SN:
                        r.setValue(TH.col_Judge, 'PASS')
                        self.txtMsg_append(rComment,Qt.blue)
                    else:                    
                        r.setValue(TH.col_Judge, 'FAIL')
                        self.txtMsg_append(rComment,Qt.red)
                MeasDut=self.DUT.getMeasure()  
                rComment+=f", Status:{MeasDut['Status']}"
                r.setValue(TH.col_Comment, rComment)
                
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
                self.Update_TRecord(r,True,rStartTime)
                sIndex = self.tbView.model().index(i, 0)
                self.tbView.scrollTo(sIndex)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort() 
        sJudge=self.JudgeResult(strIName)
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge
    
    #寫入機器日期時間    
    def Write_DUT_DateTime(self,FixUse='all'):
        subProcName="Write_DUT_DateTime"
        strIName=str("寫入機器日期時間")        
        TH=recordHeader()        
        sJudge=''                
        if self.DUT.instrConnected==0 :
            self.txtMsg_append(f"無法進行{strIName},待測機未連線",Qt.red)
            self.delay(0.5)
            return sJudge
        self.txtMsg_append(f"*** {strIName} ***",Qt.blue)        
        QtWidgets.QApplication.processEvents()
        try:
            self.DUT.openPort()        
            self.DUT.CLS()
            self.delay(0.1)
            for i in range(0,window.TRecord.model.rowCount()):
                QtWidgets.QApplication.processEvents()
                if self.StopTest: break     #若已按停止鈕, 則不再執行迴圈 
                if self.SpecTRow!=[] and not i in self.SpecTRow: continue     #有指定測試列, 但目前列, 不在指定測試列, 略過                      
                r=window.TRecord.model.record(i)
                if not fnmatch(r.value(TH.col_ItemName),strIName): continue #測試項名稱不符, 略過
                if not(FixUse=='all' or FixUse==r.value(TH.col_Cond_Fixture)): continue 
                if self.chkFixOptimize.isChecked()==False and not(r.value(TH.col_Cond_Fixture) in self.checkedFix):
                    continue                #此治具没有勾選, 執行下一回圈
                rStartTime=QDateTime.currentDateTime().toString(self.MyDTFormat)        #記錄此筆測試開始時間
                self.txtMsg_append(f"{r.value(TH.col_ItemName)},{r.value(TH.col_ItemName_2t)},{r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} {r.value(TH.col_CondUnit_1)} 確認中...")
                self.Blank_TRecord(r) 
                rComment=''
                r.setValue(TH.col_Comment, rComment)
                rindex = window.TRecord.model.index(i, 0)
                window.TRecord.model.TestRow=i
                window.TRecord.tbview.scrollTo(rindex)
                QtWidgets.QApplication.processEvents()
                self.SwitchFix( r.value(TH.col_Cond_Fixture))   #切換治具
                stm=monotonic()     #設等待的時間起點
                i2t=str(r.value(TH.col_ItemName_2t)).upper().replace(" ","")
                i3t=str(r.value(TH.col_ItemName_3t)).upper().replace(" ","")
                wDateTime_Format='yyyy_MM_dd_hh_mm_ss'
                currTime=QDateTime.currentDateTime()
                cdatetime= currTime.toString(wDateTime_Format) 
                set_datetime=cdatetime.split('_')
                self.DUT.setDateTime(set_datetime[0],set_datetime[1],set_datetime[2],set_datetime[3],set_datetime[4],set_datetime[5],0,0.5)
                get_datetime=self.DUT.getDateTime()

                rComment=f"設定日期時間:{currTime.toString('yyyy-MM-dd hh:mm:ss')}, 讀回日期時間:{get_datetime[0]}-{get_datetime[1]}-{get_datetime[2]} {get_datetime[3]}:{get_datetime[4]}:{get_datetime[5]}"
                sdvalue=int(set_datetime[0])*365+int(set_datetime[1])*30+int(set_datetime[2])
                gdvalue=int(get_datetime[0])*365+int(get_datetime[1])*30+int(get_datetime[2])
                stvalue=int(set_datetime[3])*365+int(set_datetime[4])*30+int(set_datetime[5])
                gtvalue=int(get_datetime[3])*365+int(get_datetime[4])*30+int(get_datetime[5])
                
                # r.setValue(TH.col_Comment, rComment)
                if (sdvalue==gdvalue and gtvalue>=stvalue and gtvalue<=stvalue+3):             #若日期時間相符,PASS
                    r.setValue(TH.col_Judge, 'PASS')
                    self.txtMsg_append(rComment,Qt.blue)
                else:               #若序號不相符, 燒序號, 再讀序號比對     
                    r.setValue(TH.col_Judge, 'FAIL')
                    self.txtMsg_append(rComment,Qt.red)                    
                    
                MeasDut=self.DUT.getMeasure()  
                rComment+=f", Status:{MeasDut['Status']}"
                r.setValue(TH.col_Comment, rComment)
                
                DErr=self.DUT.getSysErr()
                if DErr!='':
                    self.txtMsg_append(f"{r.value(TH.col_ItemName_2t)}, {r.value(TH.col_ItemName_3t)}, {r.value(TH.col_Cond_1)} \nDUT發生錯誤:{str(DErr)}",Qt.red)
                self.Update_TRecord(r,True,rStartTime)
                sIndex = self.tbView.model().index(i, 0)
                self.tbView.scrollTo(sIndex)
        except Exception as e:
            self.txtMsg_append(f"{subProcName} Error: \n{str(e)}",Qt.red)
        self.DUT.closePort() 
        sJudge=self.JudgeResult(strIName)
        window.TRecord.model.TestRow=-1
        try:
            window.TRecord.tbview.scrollTo(rindex)  
        except Exception as e:
            serr=e        
        return sJudge

    #Initial
    def Set_DUT_Initial(self,FixUse='all'):
        self.txtMsg.setText('Initial')


    

        

if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    try:
        window = MainWindow()    
        window.show()
        QtWidgets.QApplication.processEvents()
        window.onQApplicationStarted()
    except Exception as e:
        QMessageBox.critical(window,'啟動錯誤',f'Error:{str(e)}')
    sys.exit(app.exec_())