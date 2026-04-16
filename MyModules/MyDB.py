import sys
import sqlite3
from datetime import date, datetime
from time import sleep
import pathlib

Ver='0.2.0'

class SQLiteDB():
    ErrorMessages=''
    con=sqlite3.Connection
    #連接資料庫
    def ConnectSQLite(self, FileFullName):
        self.con = sqlite3.connect(FileFullName)

    #關閉資料庫連線
    def CloseConnection(self):    
        self.con.close()

    #取回指定名稱的資料表內容
    def getTable(self, TName="", strFilter=""):
        strQ=""
        self.con.row_factory = sqlite3.Row      #此列讓fetchall()或fetchone()函數傳回每列為sqlite3.Row, 可像dictionary一樣操作資料
        if strFilter=="":
            strQ='SELECT * FROM ' + TName
        else:
            strQ='SELECT * FROM ' + TName+ 'where ' + strFilter 
        cur=self.con.execute(strQ)
        rows=cur.fetchall()               
        return rows

    def Query(self, SQLString=''):
        self.con.row_factory = sqlite3.Row      #此列讓fetchall()或fetchone()函數傳回每列為sqlite3.Row, 可像dictionary一樣操作資料
        rows=[]
        rows.clear
        try:
            if SQLString!="":            
                cur=self.con.execute(SQLString)
                rows=cur.fetchall() 
        except Exception as e:
            print(str(e))
        return rows        

    #插入一筆記錄到資料表
    def InsertRecord(self, TName, record={}):
        SubName='InsertRecord'
        try:
            tb='Insert into '+ TName 
            fd=''
            v=''
            # print(type(record))
            for r in record:
                fd+=r+', '
                # print(type(record[r]).__name__)
                if type(record[r]).__name__=='str':
                    v+='\''+record[r]+'\', '
                else:
                    v+=''+str(record[r])+', '
            fd=fd[0:len(fd)-2]
            v=v[0:len(v)-2]
            s=tb+' ('+fd+') values ('+ v +')'
            # print(s)
            self.con.execute(s)
            self.con.commit()            
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            raise e

    def TestConnection(self):
        rs=self.getTable("FIX6501")
        print(rs[0][0], rs[0][1], rs[0][2])
        for r in rs:
            for i in range(0,len(r)):
                print(r[i])
                print()

