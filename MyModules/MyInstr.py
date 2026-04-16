#from curses import baudrate
# from asyncio.windows_events import NULL
# from tokenize import String
import pyvisa as visa
from pyvisa import constants
from time import *
from fnmatch import fnmatch
from PyQt5 import QtWidgets
# import pyvisa
import win32com.client
import inspect
from MyModules.Base_Instr import Instrument
from MyModules.MyFunctions import isnumber
#from enum import Enum

def isnumber(aString):
    try:
        float(aString)
        return True
    except:
        return False

#**********  儀器底層物件  **********
class Instrument():
    VisaDescription=""
    Baudrate=9600
    ErrorMessages=""
    instrOpened=False
    instrIDN=""
    instrTimeouts_ms=500
    instrDelayTime_s=0.1
    instrPassWord="abc123"
    instrConnected=0
    IDN_Keyword=""
    LastAnswer=''   

    rm = visa.ResourceManager()
    instr=visa.resources.Resource(rm,"")    
    def delay(self,delaySec=0.0):
        wt=0
        st=monotonic()
        while(wt<delaySec):
            wt=monotonic()-st
            QtWidgets.QApplication.processEvents()
            sleep(0.001)    
    #以關鍵字找機器的通訊埠, 找到傳回Port的Visa Description
    def findPort(self, keyword=""):
        #rm = visa.ResourceManager()
        SubName="findPort"
        foundPort="" 
        dev=""
        s=""
        if keyword=="":
            keyword=self.IDN_Keyword

        if keyword=="":
            strDev_2find="psu"      #預設找PSU
        else:
            strDev_2find=keyword.lower()
        if self.VisaDescription!="":
            if self.checkPort(self.VisaDescription,strDev_2find):
                foundPort=self.VisaDescription
        if foundPort=="":
            tupDevices=self.rm.list_resources()
            for dev in tupDevices:
                if self.checkPort(dev,strDev_2find):
                    foundPort=dev
                    break       
        
        if fnmatch(self.VisaDescription,"ASRL*"):
            self.instr = visa.resources.SerialInstrument(self.rm, self.VisaDescription)  #設定resource為序列埠儀器
        elif fnmatch(self.VisaDescription,"GPIB*"):
            self.instr =visa.resources.GPIBInstrument(self.rm, self.VisaDescription)     #設定resource為GPIB儀器
        return foundPort

    def getErrMessage(self):
        s=self.ErrorMessages
        self.ErrorMessages=''
        return s

    #檢查port傳回值是否符合要查的關鍵字
    def checkPort(self, DevVisaDescript="", keyWord="", commandString='*IDN?'):
        SubName="checkPort"
        checkOK=False
        print(f"{self.VisaDescription}")
        #rm = visa.ResourceManager()
        if DevVisaDescript=="":
            DevVisaDescript=self.VisaDescription
        if keyWord=="":
            keyWord=self.IDN_Keyword
        try:            
            i=0
            fd=False
             
            inst=self.rm.open_resource(DevVisaDescript)

            if fnmatch(DevVisaDescript,'*tcpip*'):
                inst.timeout=1000
                inst.set_visa_attribute(constants.VI_ATTR_TERMCHAR_EN,constants.VI_TRUE)
            else:
                inst.timeout=self.instrTimeouts_ms            
            if self.Baudrate!=9600:
                inst.baud_rate=self.Baudrate
            while(i<3 and fd==False):
                s=''
                try:               
                    s = inst.query(commandString)
                except Exception as e:
                    self.ErrorMessages+=f"checkPort {DevVisaDescript}, Err: {e}\n"
                s=str(s).replace('\n','').replace('\r','').strip()
                if s!='':
                    # print(f"checkPort s=>{s}")
                    fd=True
                    break
                else:
                    print(f"checkPort {DevVisaDescript},  i={i}, delay:{self.instrDelayTime_s}")
                i+=1
                self.delay(self.instrDelayTime_s)
                inst.flush(constants.VI_READ_BUF_DISCARD)     #清除先前命令留在讀回值緩衝區的值

            if s.lower().find(keyWord.lower())>=0:                    
                checkOK=True
                self.instrIDN=s         #把IDN先記起來
                self.VisaDescription=DevVisaDescript    #把Visa Description記下
            inst.close()
        except visa.errors.VisaIOError as e:
            if e.error_code == constants.StatusCode.error_timeout:                          
                self.ErrorMessages+=SubName+ " ("+DevVisaDescript+") timeout!\n"            
            else:
                self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
                #raise e                #已知找port會有timeout,resource_not_found等狀況, 忽略此錯誤.
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            raise e                
        finally:            
            return checkOK
    

    #開啟通訊埠
    def openPort(self):
        SubName="openPort"
        self.instrOpened=False
        #self.VisaDescription = "ASRL5::INSTR"
        print(self.VisaDescription)
        s=self.VisaDescription   
        if not s:  # 檢查是否為空
            self.ErrorMessages += SubName + " Error: VisaDescription is empty.\n"
            return  # 直接返回，避免後續錯誤     
        try:
            self.instr=self.rm.open_resource(self.VisaDescription)
            if s.lower().find('asrl')>=0:
                self.instr.baud_rate=self.Baudrate
                self.instr.timeout=self.instrTimeouts_ms
            elif s.lower().find('tcpip')>=0:
                self.instr.timeout=5000 
                self.instr.set_visa_attribute(constants.VI_ATTR_TERMCHAR_EN,constants.VI_TRUE)               
            else:
                self.instr.timeout=self.instrTimeouts_ms
            self.instrOpened=True
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
            #raise e       
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e

    #關閉通訊埠
    def closePort(self):
        SubName="closePort"        
        try:
            self.instr.close()
            self.instrOpened=False
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
            raise e       
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            raise e
    def get_delay_args(self,args):
        delay_before=0
        delay_after=0
        for key, value in args:   #讀取動態參數
            print(f"{key}: {value}")
            if key.lower() in ["delay_before","wait_before"]:
                delay_before=value
            elif key.lower() in ["delay","delay_after","wait_after"]:
                delay_after=value
        return delay_before,delay_after

    #送命令給機器, 再讀回回傳值
    def sendQuery(self,QueryString="",**kwargs):
        """
        送命令給機器, 再讀回回傳值
        參數:
        QueryString: string => 要傳送的查詢命令 
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        string, 機器回傳字串
        """
        SubName=inspect.currentframe().f_code.co_name 
        strAns=""
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)        #先等待指定時間,再執行
            if not(self.instrOpened):
                self.openPort()                
            self.instr.flush(constants.VI_READ_BUF_DISCARD)     #清除先前命令留在讀回值緩衝區的值
            self.instr.write_termination = '\n'
            self.instr.read_termination = '\n'
            strAns=self.instr.query(QueryString)
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=f'{SubName}, {QueryString}\nvisa.errors.VisaIOError: {e.description}\n'
            #raise e       
        except Exception as e:
            #self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            raise e
        finally:
            self.LastAnswer=strAns
            self.delay(delay_after)         #等待指定時間, 再結束
            return strAns

    #送命令給機器
    def sendCommand(self,CommandString="",**kwargs):
        """
        送命令給機器
        參數:
        CommandString: string => 要傳送的命令 
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        無
        """
        SubName=inspect.currentframe().f_code.co_name         
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)        #先等待指定時間,再執行
            if not(self.instrOpened):
                self.openPort()
            self.instr.write(CommandString)
            self.delay(delay_after)         #等待指定時間, 再結束
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=f'{SubName}, {CommandString}\nvisa.errors.VisaIOError: {e.description}\n'
            #raise e       
        except Exception as e:
            self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            #raise e

    #讀回機器傳回訊息
    def readString(self):
        SubName="readString"
        strAns=""         
        try:
            strAns=self.instr.read()
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
            #raise e       
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        finally:
            return strAns
    #讀回IDN
    def get_IDN(self):
        SubName='get_IDN'
        m=0
        fd=False 
        while(m<3 and fd==False):            
            id=''
            try:
                id=str(self.sendQuery('*IDN?')).replace('\r','').replace('\n','').strip()
                print(id)
            except Exception as e:
                self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            if id!='':
                fd=True 
                break                   
            if m>1:
                print(f"{self.VisaDescription} get_IDN m=>{m}")                    
            m+=1
            self.delay(self.instrDelayTime_s)
        return id 

    def CLS(self):
        self.sendCommand('*CLS')

    def RST(self):
        self.sendCommand('*RST')

    def getDevInfo(self):
        dinf={}
        try:
            # usb.DeviceID: 返回设备的唯一标识符。
            # usb.Name: 返回设备的名称，通常包含COM端口号。
            # usb.Description: 返回设备的描述信息。
            # usb.Caption: 返回设备的标题或标签。            
            # usb.PNPDeviceID: 返回设备的即插即用设备ID。
            # usb.Status: 返回设备的当前状态，如"OK"或"Error"。
            # usb.ProviderType: 返回提供程序类型，通常为"Modem"。
            # usb.SystemName: 返回设备所连接的计算机名称。
            # usb.CreationClassName: 返回表示设备类别的字符串。
            
            wmi = win32com.client.GetObject ("winmgmts:")
            for usb in wmi.InstancesOf ("Win32_SerialPort"):
                try:
                    if usb.DeviceID==self.instr.resource_info.alias:
                        dinf["Device ID"]= usb.DeviceID
                        dinf["Name"]= usb.name
                        dinf["System Name"]= usb.SystemName
                        dinf["Caption"]= usb.Caption                                          
                        dinf["Description"]= usb.Description
                        dinf["PNPDeviceID"]= usb.PNPDeviceID
                        dinf["Status"]= usb.Status
                        

                except Exception as e:
                    print(str(e))
        except Exception as e:
            print("Error => "+str(e)+"\n")
        return dinf
    
#**********  DMM物件  **********
class DMM(Instrument):
    #取得主顯的讀值
    def getMainReading(self,**kwargs):
        SubName="getMainReading"
        s=''
        rdg='0.0'
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)        #先等待指定時間,再執行
            s=self.sendQuery("VAL1?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            try:
                rdg =float(s)
            except ValueError:
                self.ErrorMessages += f"{SubName} invalid value: {s}\n"
 
            self.delay(delay_after)         #等待指定時間, 再結束
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        return rdg
    #設定主顯為DCV, 檔位
    def setDCV(self,range="auto"):
        if range.lower()=="auto":
            self.sendCommand("CONFigure:VOLTage:DC")
            self.sendCommand("CONFigure:AUTO ON")
        else:
            self.sendCommand("CONFigure:VOLTage:DC "+range)

    #設定主顯為ACV, 檔位
    def setACV(self,range="auto"):
        if range.lower()=="auto":
            self.sendCommand("CONFigure:VOLTage:AC")
            self.sendCommand("CONFigure:AUTO ON")
        else:
            self.sendCommand("CONFigure:VOLTage:AC "+range)
    #設定主顯為DCI, 檔位
    def setDCI(self,range="auto"):
        if range.lower()=="auto":
            self.sendCommand("CONFigure:CURRent:DC")
            self.sendCommand("CONFigure:AUTO ON")
        else:
            self.sendCommand("CONFigure:CURRent:DC "+range)
    #設定主顯為ACI, 檔位
    def setACI(self,range="auto"):
        if range.lower()=="auto":
            self.sendCommand("CONFigure:CURRent:AC")
            self.sendCommand("CONFigure:AUTO ON")
        else:
            self.sendCommand("CONFigure:CURRent:AC "+range)
    
    #設定更新率
    def setRefleshRate(self, speed="S"):
        s=speed.lower().strip()
        if s=="s" or s=="slow":
            self.sendCommand(":SENS:DET:RATE S")
        elif s=="m" or s=="midium":
            self.sendCommand(":SENS:DET:RATE M")
        elif s=="f" or s=="fast":
            self.sendCommand(":SENS:DET:RATE F")

    
    #設定ACI顯示解析度, MIN位數最多, 更新速度慢, MAX位數最少, 更新速度快
    def setACIRESolution(self, res='MIN'):        
        self.sendCommand(f":SENSe:CURRent:AC:RESolution {res}")
    #設定DCI顯示解析度, MIN位數最多, 更新速度慢, MAX位數最少, 更新速度快
    def setDCIRESolution(self, res='MIN'):        
        self.sendCommand(f":SENSe:CURRent:DC:RESolution {res}")

    #設定Calculate State
    def setCalculateState(self, OnOff="on"):
        s=""
        if type(OnOff)==str:
            s=OnOff.lower().strip()
        elif type(OnOff)==int:
            s=str(OnOff)
        if s=="on" or s=="1":
            self.sendCommand("CALCulate:STATE ON")
        elif s=="off" or s=="0":
            self.sendCommand("CALCulate:STATE OFF")

    #設定Calculate Function
    def setCalculateFunction(self, CalFunction):
        s=CalFunction.lower().strip()
        if s=="off":
            self.sendCommand("CALCulate:FUNCtion OFF")
        elif s=="min":
            self.sendCommand("CALCulate:FUNCtion MIN")
        elif s=="max":
            self.sendCommand("CALCulate:FUNCtion MAX")
        elif s=="hold":
            self.sendCommand("CALCulate:FUNCtion HOLD")
        elif s=="rel":
            self.sendCommand("CALCulate:FUNCtion REL")
        elif s=="comp":
            self.sendCommand("CALCulate:FUNCtion COMP")
        elif s=="db":
            self.sendCommand("CALCulate:FUNCtion DB")
        elif s=="dbm":
            self.sendCommand("CALCulate:FUNCtion DBM")
        elif s=="store":
            self.sendCommand("CALCulate:FUNCtion STORE")
        elif s=="aver":
            self.sendCommand("CALCulate:FUNCtion AVER")
        elif s=="mxb":
            self.sendCommand("CALCulate:FUNCtion MXB")
        elif s=="inv":
            self.sendCommand("CALCulate:FUNCtion INV")
        elif s=="ref":
            self.sendCommand("CALCulate:FUNCtion REF")

    #讀回Calculate Max值
    def getCalculateMax(self):        
        SubName="getCalculateMax"
        s=''
        rdg=''
        try:
            s=self.sendQuery("CALCulate:MAXimun?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        return rdg

    #讀回Calculate Min值
    def getCalculateMin(self):       

        SubName="getCalculateMin"
        s=''
        rdg=''
        try:
            s=self.sendQuery("CALCulate:MINimun?")
            s=str(s).strip().replace('\n','').replace('\r','') 
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        return rdg   

#**********  PCS-1000物件  **********
class PCS1000(Instrument):
    #設定主顯為DCI, 檔位
    def setDCI(self,range="auto",**kwargs):
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)
            if type(range)==float or type(range)==int:
                self.sendCommand("CONFigure:CURRent:DC "+str(range))
                self.delay(delay_after)
            elif range.lower()=="auto":
                self.sendCommand("CONFigure:CURRent:DC AUTO")
                self.delay(delay_after)            
            else:
                self.sendCommand("CONFigure:CURRent:DC "+range)
                self.delay(delay_after)
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
            raise e       
        except Exception as e:
            self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
            raise e

    #設定主顯為ACI, 檔位
    def setACI(self,range="auto",**kwargs):
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before) 
            if type(range)==float or type(range)==int:
                self.sendCommand("CONFigure:CURRent:AC "+str(range))
                self.delay(delay_after)
            elif range.lower()=="auto":
                self.sendCommand("CONFigure:CURRent:AC AUTO")
                self.delay(delay_after)            
            else:
                self.sendCommand("CONFigure:CURRent:AC "+range)
                self.delay(delay_after)
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
            raise e       
        except Exception as e:
            self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
            raise e
    #讀回DCI的值
    def getReading_DCI(self,**kwargs):
        # s=self.sendQuery("MEASure:CURRent:DC?")
        # return s
        SubName="getReading_DCI"
        s=''
        rdg=''
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)
            s=self.sendQuery("MEASure:CURRent:DC?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            self.delay(delay_after)         #等待指定時間, 再結束
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        return rdg

    #讀回ACI的值
    def getReading_ACI(self,**kwargs):
        # s=self.sendQuery("MEASure:CURRent:AC?")
        # return s
        SubName="getReading_ACI"
        s=''
        rdg=''
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)
            s=self.sendQuery("MEASure:CURRent:AC?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            self.delay(delay_after)         #等待指定時間, 再結束
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        return rdg

    def getReading_DCV(self,**kwargs):
        # s=self.sendQuery("MEASure:CURRent:AC?")
        # return s
        SubName="getReading_DCV"
        s=''
        rdg=''
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)
            s=self.sendQuery("MEASure:VOLT:DC?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            self.delay(delay_after)         #等待指定時間, 再結束
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
            #raise e
        return rdg
#**********  Vitrek 4700物件  **********        
class vitrek_4700(Instrument):    
    def getDCV(self):        
        SubName="getDCV"
        s=''
        rdg=''
        try:
            s=self.sendQuery("DCV?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"            
        return rdg

    def getACV(self):        
        SubName="getACV"
        s=''
        rdg=''
        try:
            s=self.sendQuery("ACV?") 
            s=str(s).strip().replace('\n','').replace('\r','')
            if isnumber(s):
                rdg=float(s)
            else:
                rdg=''        
        except Exception as e:
            self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"            
        return rdg


#**********  GPT物件  **********
class GPT(Instrument):    
    TFunction=""
    TJudgment=""
    Test_voltage=0
    Test_currentOrResistance=0
    Test_time=0     

    def __init__(self):
        self.EditMode=''

    #設定是否接地    
    def setGroundMode(self,OnOff,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        if str(OnOff).lower() == 'off' or OnOff==0:
            self.sendCommand("MANU:UTILity:GROUNDMODE OFF")
        else:
            self.sendCommand("MANU:UTILity:GROUNDMODE ON")
        if delayAf>0 : self.delay(delayAf)
    
    def setACW_Freq_Volt(self,FreqHz, VoltKV,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        if str(FreqHz) in ['50','60']:
            self.sendCommand(f"MANU:ACW:FREQuency {FreqHz}")
        self.delay(0.05)
        self.sendCommand(f"MANU:ACW:VOLTage {VoltKV}")
        if delayAf>0 : self.delay(delayAf)

    def setACW_Freq(self,FreqHz,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        if str(FreqHz) in ['50','60']:
            self.sendCommand(f"MANU:ACW:FREQuency {FreqHz}")
        if delayAf>0 : self.delay(delayAf)

    def setACW_Volt(self,VoltKV,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:ACW:VOLTage {VoltKV}")
        if delayAf>0 : self.delay(delayAf)

    def setACW_Timer(self,sec=1,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:ACW:TTIMe {sec}")
        if delayAf>0 : self.delay(delayAf)

    def setACW_HiSet(self,mA,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:ACW:CHISet {mA}")
        if delayAf>0 : self.delay(delayAf)
    
    def setACW_LoSet(self,mA,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:ACW:CLOSet {mA}")
        if delayAf>0 : self.delay(delayAf)

    def setDCW_Volt(self,VoltKV,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:DCW:VOLTage {VoltKV}")
        if delayAf>0 : self.delay(delayAf)

    def setDCW_Timer(self,sec=1,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:DCW:TTIMe {sec}")
        if delayAf>0 : self.delay(delayAf)
    
    def setDCW_HiSet(self,mA,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:DCW:CHISet {mA}")
        if delayAf>0 : self.delay(delayAf)
    
    def setDCW_LoSet(self,mA,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:DCW:CLOSet {mA}")
        if delayAf>0 : self.delay(delayAf)

    def setIR_Volt(self,VoltKV,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:IR:VOLTage {VoltKV}")
        if delayAf>0 : self.delay(delayAf)

    def setIR_Timer(self,sec=1,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:IR:TTIMe {sec}")
        if delayAf>0 : self.delay(delayAf)
    
    def setIR_HiSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        if str(MOHM).upper()=='NULL':
            self.sendCommand(f"MANU:IR:RHISet {MOHM}")
        elif isnumber(MOHM):
            if float(MOHM)>9999:
                MOHM=9999
            self.sendCommand(f"MANU:IR:RHISet {MOHM}")        
        if delayAf>0 : self.delay(delayAf)
    
    def setIR_LoSet(self,MOHM,delayBe=0, delayAf=0):        
        if delayBe>0 : self.delay(delayBe)        
        if isnumber(MOHM):
            M=float(MOHM)
            if M>9999:
                MOHM=9999
            elif M<1:
                MOHM=1
        self.sendCommand(f"MANU:IR:RLOSet {MOHM}")        
        if delayAf>0 : self.delay(delayAf)
    
    def getIR_LoSet(self,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        r=self.sendQuery(f"MANU:IR:RLOSet?")
        if delayAf>0 : self.delay(delayAf)
        return r

    def setGB_Curr(self,CurrentA,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:GB:Current {CurrentA}")
        if delayAf>0 : self.delay(delayAf)
    
    def setGB_Freq(self,FreqHz,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:GB:FREQuency {FreqHz}")
        if delayAf>0 : self.delay(delayAf)

    def setGB_Timer(self,sec=1,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:GB:TTIMe {sec}")
        if delayAf>0 : self.delay(delayAf)
    
    def setGB_HiSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:GB:RHISet {MOHM}")
        if delayAf>0 : self.delay(delayAf)
    
    def setGB_LoSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:GB:RLOSet {MOHM}")
        if delayAf>0 : self.delay(delayAf)

    def setEditMode(self, editMode='ACW',delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:EDIT:MODE {editMode}")
        self.EditMode=editMode
        if delayAf>0 : self.delay(delayAf)
    
    def getEditMode(self,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        r=self.sendQuery(f"MANU:EDIT:MODE?")
        self.EditMode=r
        if delayAf>0 : self.delay(delayAf)
        return r

    def setRampTime(self, sec=0.1,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:RTIMe {sec}")
        if delayAf>0 : self.delay(delayAf) 

    def getRampTime(self,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        r=self.sendQuery(f"MANU:RTIMe?")
        if delayAf>0 : self.delay(delayAf)
        return r
    #傳回字串, 從數字開始, 並除去多的前置0
    def del_Pre_0(self,s):
        len_S=len(s)
        for j in range(0,len_S):
            if s[j]=='0' and j<(len_S-1):
                if s[j+1]==' ' or s[j+1]=='.' or s[j+1]>='A':
                    break
            if (s[j]>='1' and s[j]<='9') or s[j]=='.':
                if s[j]=='.':
                    j-=1
                break
        s=s[j:]
        return s

    def getMeasure_RawStr(self,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        try:
            q=''
            q=self.sendQuery("MEASure?")
        except:
            pass
        return q
    def getMeasure(self,delayBe=0, delayAf=0):
        meas={"Function":"", "Status":"", "TVA":"", "TAR":"", "TT":""}
        if delayBe>0 : self.delay(delayBe)
        q=self.sendQuery("MEASure?")
        # print(f"getMeasure回傳=>{q}")
        sl=str(q).split(',')
        for i in range(len(sl)):
            sl[i]=sl[i].strip()
        try:
            meas["Function"]=sl[0]
        except:
            meas["Function"]=""
        try:
            if sl[1].lower().find('fail')>=0:
                meas["Status"]='FAIL'
            else:
                meas["Status"]=sl[1]
        except:
            meas["Status"]=""
        
        try:
            sl[2]=sl[2].lower().replace('kv','').replace('m','').replace('a','').strip()
            for j in range(0,len(sl[2])):
                if (sl[2][j]>='1' and sl[2][j]<='9') or sl[2][j]=='.':
                    if sl[2][j]=='.':
                        j-=1
                    break
            sl[2]=sl[2][j:]
            if isnumber(sl[2]):
                meas["TVA"]=float(sl[2])
        except:
            meas["TVA"]=""
        try:
            sl_3=sl[3].lower()
            sl[3]=sl[3].lower().replace('ohm','').replace('m','').replace('a','').strip()
            
            # for j in range(0,len(sl[3])):
            #     if (sl[3][j]>='1' and sl[3][j]<='9') or sl[3][j]=='.':
            #         if sl[3][j]=='.':
            #             j-=1
            #         break
            # sl[3]=sl[3][j:]
            sl[3]=self.del_Pre_0(sl[3])
            
            if isnumber(sl[3]):
                meas["TAR"]=float(sl[3])
            elif isnumber(sl[3].lower().replace('g','')):
                meas["TAR"]=float(sl[3].lower().replace('g',''))*1000
            elif str(sl[3]).lower().find('u')>0:
                a=sl[3].lower().replace('u','').replace('a','').strip()
                if a=='000' or a=='00' or a=='000.0' or a=='00.0':
                    a=0
                if isnumber(a):                
                    meas["TAR"]=float(a)/1000   #uA轉mA
        except:
            meas["TAR"]=""

        try:
            meas["TT"]=sl[4]
        except:
             meas["TT"]=""
        
        if delayAf>0 : self.delay(delayAf)
        # print(str(meas))
        return meas
    
    def getSysErr(self, delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        q=self.sendQuery("SYST:ERR?")
        # sErr=str(q).split(",")
        sErr=str(q).strip()
        self.sendCommand("*CLS")
        # for i in range(0,len(sErr)):
        #     sErr[i]=str(sErr[i]).strip()
        if sErr.lower().find("no error"):
            sErr=""
        if delayAf>0 : self.delay(delayAf)
        return sErr

    def Test(self, OnOff,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        if str(OnOff).lower() == 'off' or OnOff==0:
            self.sendCommand("FUNCtion:TEST OFF")
        elif str(OnOff).lower() == 'on' or OnOff==1:
            self.sendCommand("FUNCtion:TEST ON")
        if delayAf>0 : self.delay(delayAf)

    def writeSN(self, SN='', delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        SN=SN.strip()        
        self.sendCommand(f"SYSTem:PID:SETTing {SN}")
        if delayAf>0 : self.delay(delayAf)

    def getSN(self, delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        q=self.sendQuery(f"SYSTem:PID:SETTing?")
        q=str(q).replace('\n','').replace('\r','').strip()        
        if delayAf>0 : self.delay(delayAf)
        return q
    


    #指令有問題
    # def setInitial(self, delayBe=0, delayAf=0):
    #     if delayBe>0 : self.delay(delayBe)
    #     self.sendCommand("MANU:INITial")
    #     if delayAf>0 : self.delay(delayAf)
        
#**********  GPT99物件  **********
class GPT99(GPT):
    def setIR_HiSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)        
        if str(MOHM).upper()=='NULL':
            self.sendCommand(f"MANU:IR:RHISet {MOHM}")
        elif isnumber(MOHM):
            if float(MOHM)>50000:
                MOHM=50000
            self.sendCommand(f"MANU:IR:RHISet {MOHM}M")
        if delayAf>0 : self.delay(delayAf)
    
    def setIR_LoSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)        
        if isnumber(MOHM):
            M=float(MOHM)
            if M>50000:
                MOHM=50000
            elif M<1:
                MOHM=1
            self.sendCommand(f"MANU:IR:RLOSet {MOHM}M")
        else:
            self.sendCommand(f"MANU:IR:RLOSet {MOHM}")
        if delayAf>0 : self.delay(delayAf)

#**********  GPT10000物件  **********
class GPT10000(GPT):
    def setIR_HiSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)        
        if str(MOHM).upper()=='NULL':
            self.sendCommand(f"MANU:IR:RHISet {MOHM}")
        elif isnumber(MOHM):
            M=float(MOHM)
            if M<1000:                
                self.sendCommand(f"MANU:IR:RHISet {MOHM}M")
            elif M>=1000 and M<10000:
                self.sendCommand(f"MANU:IR:RHISet {round(M/1000,3)}G")
            elif M>=10000 and M<=50000:
                self.sendCommand(f"MANU:IR:RHISet {round(M/1000,2)}G")
            else:
                M=50000
                self.sendCommand(f"MANU:IR:RHISet {round(M/1000,2)}G")
        if delayAf>0 : self.delay(delayAf)
    
    def setIR_LoSet(self,MOHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)        
        if isnumber(MOHM):
            M=float(MOHM)
            if M<1000:                
                self.sendCommand(f"MANU:IR:RLOSet {MOHM}M")
            elif M>=1000 and M<10000:
                self.sendCommand(f"MANU:IR:RLOSet {round(M/1000,3)}G")
            elif M>=10000 and M<=50000:
                self.sendCommand(f"MANU:IR:RLOSet {round(M/1000,2)}G")
            else:
                M=50000
                self.sendCommand(f"MANU:IR:RLOSet {round(M/1000,2)}G")
        if delayAf>0 : self.delay(delayAf)

    
    #設定是否接地    
    def setGroundMode(self,OnOff,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        M=str(self.EditMode).strip().upper()        
        if str(OnOff).lower() == 'off' or OnOff==0:
            self.sendCommand(f"MANU:{M}:GROUNDMODE OFF")
        else:
            self.sendCommand(f"MANU:{M}:GROUNDMODE ON")
        if delayAf>0 : self.delay(delayAf)

    
    def setCONT_Timer(self,sec=1,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:CONTinuity:TTIMe {sec}")
        if delayAf>0 : self.delay(delayAf)
    
    def setCONT_HiSet(self,OHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:CONTinuity:RHISet {OHM}")
        if delayAf>0 : self.delay(delayAf)
    
    def setCONT_LoSet(self,OHM,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        self.sendCommand(f"MANU:CONTinuity:RLOSet {OHM}")
        if delayAf>0 : self.delay(delayAf) 
    
    #設定日期,時間
    def setDateTime(self, year, month, day, hour, minute, second,delayBe=0, delayAf=0):
        if (isnumber(year) and isnumber(month) and isnumber(day) 
            and isnumber(hour) and isnumber(minute) and isnumber(second)):   
            if delayBe>0 : self.delay(delayBe)
            strCmd=f"SYSTEM:TIME T{year[2:]}_{month}_{day}_{hour}_{minute}_{second}"
            self.sendCommand(strCmd)
            if delayAf>0 : self.delay(delayAf)
    
    #查詢日期,時間
    def getDateTime(self,delayBe=0, delayAf=0):
        if delayBe>0 : self.delay(delayBe)
        q=self.sendQuery(f"SYSTem:TIMe?")
        q=str(q).replace('T','').replace('\n','').replace('\r','').strip()
        dt=q.split(' ')
        year='1900'
        month='1'
        day='1'
        hour='0'
        minute='0'
        second='0'
        if len(dt)>=2:
            sd=dt[0]
            st=dt[1]
            d=sd.replace(' ','').split('-')
            t=st.replace(' ','').split(':')
            if len(d)>=3 and len(t)>=3:
                year=d[0]
                month=d[1]
                day=d[2]                                
                hour=t[0]
                minute=t[1]
                second=t[2]
        
        if delayAf>0 : self.delay(delayAf)
        return year, month, day, hour, minute, second

#**********  GDS1000B物件  **********
class GDS1000B(Instrument):   
    def setAverage(self,Avg):
        self.sendCommand(":ACQuire:AVERage "+str(Avg))

#**********  PEL物件  **********
class PEL(Instrument):
    class Enum_Input_State():
        OFF=0                   #input Off (load off)
        ON=1                    #input on (load on)
    
    #列舉輸入短路狀態
    class Enum_Input_Short_State():
        OFF=0                   #input short Off
        ON=1                    #input short on
    
    #列舉操作模式CC,CR,CV,CP...
    class Enum_Oper_Mode():
        CC='CC' #CC Mode
        CR='CR' #CR Mode
        CV='CV' #CV Mode
        CP='CP' #CP Mode
        CCCV='CCCV' #CC+CV Mode
        CRCV='CRCV' #CR+CV Mode
        CPCV='CPCV' #CP+CV Mode

    #列舉操作模式CC,CR,CV,CP...
    class Enum_Dynamic_Mode():
        Dynamic='Dynamic'   #Dynamic Mode
        Static='Static'     #Static Mode        
    
    class POWER_RESPONSE():
        FAST='FAST'
        NORM='NORM'
        SLOW='SLOW'

    #列舉電流檔位
    class Enum_Current_Range():
        Low='LOW'       #Low range
        Middle='MIDDLE' #Middle range
        High='HIGH'     #High range
    
    #列舉電壓檔位
    class Enum_Voltage_Range():
        Low='LOW' #Low range
        High='HIGH' #High range
        

    def __init__(self,**kwargs):
        super().__init__(**kwargs)
    
    #設定機器輸入ON或OFF
    def setInput_State(self,OnOff,**kwargs):
        """
        設定機器輸入ON或OFF (Load ON/OFF)
        參數:
        OnOff: 1, on => 輸入ON, 0, off => 輸入OFF
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)        #先等待指定時間,再執行
            s=str(OnOff).lower().strip()
            strcmd=None
            if s in [self.Enum_Input_State.ON, 'on','1']:
                strcmd=f':INPut ON'
            elif s in [self.Enum_Input_State.OFF, 'off','0']:
                strcmd=f':INPut OFF'
            if strcmd is not None:
                self.sendCommand(strcmd,**kwargs) 
                self.delay(delay_after)         #等待指定時間, 再結束  
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
            raise e       
        except Exception as e:
            self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
            raise e

    #讀機器輸入ON或OFF
    def getInput_State(self,**kwargs):
        """
        讀取機器輸入狀態 (Load ON/OFF)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        integer, 1 => 輸入ON, 0 => 輸入OFF
        """ 
        s=self.sendQuery(f':INPut?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if s in ['1','ON']:
            return self.Enum_Input_State.ON
        elif s in ['0','OFF']:
            return self.Enum_Input_State.OFF
        else:return None 

    #設定機器輸入短路ON或OFF
    def setInput_Short_State(self,OnOff,**kwargs):
        """
        設定機器輸入短路ON或OFF (Load ON/OFF)
        參數:
        OnOff: 1, on => 輸入ON, 0, off => 輸入OFF
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(OnOff).lower().strip()
        strcmd=None
        if s in [str(self.Enum_Input_State.ON).upper(), 'on','1']:
            strcmd=f':INPut:SHORt ON'
        elif s in [str(self.Enum_Input_State.OFF).upper(), 'off','0']:
            strcmd=f':INPut:SHORt OFF'
        if strcmd is not None:
            self.sendCommand(strcmd,**kwargs)   

    #讀機器輸入短路ON或OFF
    def getInput_Short_State(self,**kwargs):
        """
        讀取機器輸入短路狀態 (Load ON/OFF)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        integer, 1 => 輸入短路ON, 0 => 輸入短路OFF
        """ 
        s=self.sendQuery(f':INPut:SHORt?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if s in ['1','ON']:
            return self.Enum_Input_Short_State.ON
        elif s in ['0','OFF']:
            return self.Enum_Input_Short_State.OFF
        else:return None 

    #設定機器負載操作模式
    def setOper_Mode(self,OP_Mode,**kwargs):
        """
        設定機器負載操作模式
        參數:
        OP_Mode string: CC => CC Mode, CR => CR Mode, CV => CV Mode, 
                CCCV => CC+CV Mode, CRCV => CR+CV Mode, CPCV => CP+CV Mode
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(OP_Mode).upper().strip()
        s1=None
        if s in [str(self.Enum_Oper_Mode.CC).upper(), 'CC']:
            s1=f'CC'
        elif s in [str(self.Enum_Oper_Mode.CR).upper(), 'CR']:
            s1=f'CR'
        elif s in [str(self.Enum_Oper_Mode.CV).upper(), 'CV']:
            s1=f'CV'
        elif s in [str(self.Enum_Oper_Mode.CP).upper(), 'CP']:
            s1=f'CP'
        elif s in [str(self.Enum_Oper_Mode.CCCV).upper(), 'CCCV']:
            s1=f'CCCV'
        elif s in [str(self.Enum_Oper_Mode.CRCV).upper(), 'CRCV']:
            s1=f'CRCV'
        elif s in [str(self.Enum_Oper_Mode.CPCV).upper(), 'CPCV']:
            s1=f'CPCV'
        if s1 is not None:
            self.sendCommand(f':MODE {s1}',**kwargs)   

    #讀機器負載操作模式
    def getOper_Mode(self,**kwargs):
        """
        讀取機器負載操作模式
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        string, CC => CC Mode, CR => CR Mode, CV => CV Mode, 
                CCCV => CC+CV Mode, CRCV => CR+CV Mode, CPCV => CP+CV Mode
        """ 
        s=self.sendQuery(f':MODE?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if s in [self.Enum_Oper_Mode.CC,'CC']:
            return self.Enum_Oper_Mode.CC            
        elif s in [self.Enum_Oper_Mode.CR,'CR']:
            return self.Enum_Oper_Mode.CR
        elif s in [self.Enum_Oper_Mode.CV,'CV']:
            return self.Enum_Oper_Mode.CV
        elif s in [self.Enum_Oper_Mode.CP,'CP']:
            return self.Enum_Oper_Mode.CP
        elif s in [self.Enum_Oper_Mode.CCCV,'CCCV']:
            return self.Enum_Oper_Mode.CCCV
        elif s in [self.Enum_Oper_Mode.CRCV,'CRCV']:
            return self.Enum_Oper_Mode.CRCV
        elif s in [self.Enum_Oper_Mode.CPCV,'CPCV']:
            return self.Enum_Oper_Mode.CPCV
        else:return None

    #設定機器電流操作檔位
    def setCurrent_Range(self,Cur_Range,**kwargs):
        """
        設定機器電流操作檔位
        參數:
        Cur_Range string: low => Low Range, High => High Range
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(Cur_Range).upper().strip()
        s1=None
        if s in [str(self.Enum_Current_Range.Low).upper(), 'LOW']:
            s1=f'LOW'
        elif s in [str(self.Enum_Current_Range.High).upper(), 'HIGH']:
            s1=f'HIGH'
        elif s in [str(self.Enum_Current_Range.Middle).upper(), 'MIDDLE','MID']:
            s1=f'MIDD'
        if s1 is not None:
            self.sendCommand(f':CRANge {s1}',**kwargs)   

    #讀機器電流操作檔位
    def getCurrent_Range(self,**kwargs):
        """
        讀取機器電流操作檔位
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        string, Low => Low Range, High => High Range
        """ 
        s=self.sendQuery(f':CRANge?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if s in [str(self.Enum_Current_Range.Low).upper(),'LOW']:
            return self.Enum_Current_Range.Low            
        elif s in [str(self.Enum_Current_Range.High).upper(),'HIGH']:
            return self.Enum_Current_Range.High    
        elif s in [str(self.Enum_Current_Range.Middle).upper(),'MIDDLE','MID']:
            return self.Enum_Current_Range.Middle    
        else:return None

    #設定機器電壓操作檔位
    def setVoltage_Range(self,Volt_Range,**kwargs):
        """
        設定機器電壓操作檔位
        參數:
        Volt_Range string: low => Low Range, High => High Range
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(Volt_Range).upper().strip()
        s1=None
        if s in [self.Enum_Voltage_Range.Low, 'LOW']:
            s1=f'LOW'
        elif s in [self.Enum_Voltage_Range.High, 'HIGH']:
            s1=f'HIGH'
        if s1 is not None:
            self.sendCommand(f':VRANge {s1}',**kwargs)   

    #讀取機器電壓操作檔位
    def getVoltage_Range(self,**kwargs):
        """
        讀取機器電壓操作檔位
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        string, Low => Low Range, High => High Range
        """ 
        s=self.sendQuery(f':VRANge?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if s in [self.Enum_Voltage_Range.Low,'LOW']:
            return self.Enum_Voltage_Range.Low            
        elif s in [self.Enum_Voltage_Range.High,'HIGH']:
            return self.Enum_Voltage_Range.High      
        else:return None

    #設定機器CC電流(A Value)
    def setCC_Current(self,A,**kwargs):
        """
        設定機器機器CC電流(A Value)
        參數:
        A:float => 電流A ,MAX => 最大可設定電流, MIN => 最小可設定電流
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(A).upper().strip()        
        self.sendCommand(f':CURRent:VA {s}',**kwargs)   

    #讀取機器機器CC電流(A Value)
    def getCC_Current(self,**kwargs):
        """
        讀取機器機器CC電流(A Value)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電流A
        """ 
        s=self.sendQuery(f':CURRent:VA?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None 
    
    #設定機器CC電流(B Value)
    def setCC_Current_B_Value(self,A,**kwargs):
        """
        設定機器機器CC電流(B Value)
        參數:
        A:float => 電流A ,MAX => 最大可設定電流, MIN => 最小可設定電流
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(A).upper().strip()        
        self.sendCommand(f':CURRent:VB {s}',**kwargs)   

    #讀取機器CC電流(B Value)
    def getCC_Current_B_Value(self,**kwargs):
        """
        讀取機器機器CC電流(B Value)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電流A
        """ 
        s=self.sendQuery(f':CURRent:VB?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None   
    
    #設定機器CV電壓(A Value)
    def setCV_Voltage(self,V,**kwargs):
        """
        設定機器機器CV電壓(A Value)
        參數:
        V:float => 電壓V ,MAX => 最大可設定電壓, MIN => 最小可設定電壓
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(V).upper().strip()        
        self.sendCommand(f':VOLTage:VA {s}',**kwargs)   

    #讀取機器CV電壓(A Value)
    def getCV_Voltage(self,**kwargs):
        """
        讀取機器機器CV電壓(A Value)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電壓V
        """ 
        s=self.sendQuery(f':VOLTage:VA?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CV電壓(B Value)
    def setCV_Voltage_B_Value(self,V,**kwargs):
        """
        設定機器機器CV電壓(B Value)
        參數:
        V:float => 電壓V ,MAX => 最大可設定電壓, MIN => 最小可設定電壓
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(V).upper().strip()        
        self.sendCommand(f':VOLTage:VB {s}',**kwargs)   

    #讀取機器CV電壓(B Value)
    def getCV_Voltage_B_Value(self,**kwargs):
        """
        讀取機器機器CV電壓(B Value)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電壓V
        """ 
        s=self.sendQuery(f':VOLTage:VB?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CR電阻(A Value)
    def setCR_Resistance(self,R,**kwargs):
        """
        設定機器機器CR電阻(A Value)
        參數:
        R:float => 電阻ohm ,MAX => 最大可設定電阻, MIN => 最小可設定電阻
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(R).upper().strip()        
        self.sendCommand(f':RESistance:VA {s}',**kwargs)   

    #讀取機器CR電阻(A Value)
    def getCR_Resistance(self,**kwargs):
        """
        讀取機器機器CR電阻(A Value)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電阻ohm
        """ 
        s=self.sendQuery(f':RESistance:VA?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CR電阻(B Value)
    def setCR_Resistance_B_Value(self,R,**kwargs):
        """
        設定機器機器CR電阻(B Value)
        參數:
        R:float => 電阻ohm ,MAX => 最大可設定電阻, MIN => 最小可設定電阻
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(R).upper().strip()        
        self.sendCommand(f':RESistance:VB {s}',**kwargs)   

    #讀取機器CR電阻(B Value)
    def getCR_Resistance_B_Value(self,**kwargs):
        """
        讀取機器機器CR電阻(B Value)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電阻ohm
        """ 
        s=self.sendQuery(f':RESistance:VB?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CR電阻SlewRate
    def setCR_Resistance_RISE_Value(self,R,**kwargs):
        """
        設定機器機器CR電阻SlewRate
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電阻ohm
        """ 
        s=str(R).upper().strip()        
        self.sendCommand(f':RES:SRAT {s}',**kwargs) 

    #設定機器CP功率(A Value)
    def setCP_Power(self,W,**kwargs):
        """
        設定機器機器CP功率(A Value)
        參數:
        R:float => 功率W ,MAX => 最大可設定功率, MIN => 最小可設定功率
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(W).upper().strip()        
        self.sendCommand(f':POW:VA {s}',**kwargs)

    #讀取機器CP功率(A Value)
    def getCP_Power(self,**kwargs):
        """
        讀取機器機器CP功率(A Value)
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=self.sendQuery(f':POW:VA?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CP功率(B Value)
    def setCP_Power_B_Value(self,W,**kwargs):
        """
        設定機器機器CP功率(B Value)
        參數:
        R:float => 功率W ,MAX => 最大可設定功率, MIN => 最小可設定功率
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(W).upper().strip()        
        self.sendCommand(f':POW:VB {s}',**kwargs)

    #讀取機器CP功率(B Value)
    def getCP_Power_B_Value(self,**kwargs):
        """
        讀取機器機器CP功率(B Value)
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=self.sendQuery(f':POW:VB?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器POWER RESPONSE mode
    def setPower_response_Mode(self,R,**kwargs):
        s=str(R).upper().strip()        
        self.sendCommand(f':POW:RESP {s}',**kwargs)

    #設定機器Dynamic或Static mode
    def setDynamic_Mode(self,Dynamic_Mode,**kwargs):
        """
        設定機器Dynamic或Static mode
        參數:
        Dynamic_Mode string: Dynamic => Dynamic mode, Static => Static mode
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(Dynamic_Mode).lower().strip()
        s1=None
        if s in [str(self.Enum_Dynamic_Mode.Dynamic).lower(), 'dynamic','dyn']:
            s1=f'DYNamic'
        elif s in [str(self.Enum_Dynamic_Mode.Dynamic).lower(), 'static','stat']:
            s1=f'STATic'
        if s1 is not None:
            self.sendCommand(f':MODE:DYNamic {s1}',**kwargs)


    #讀取機器Dynamic或Static mode
    def getDynamic_Mode(self,**kwargs):
        """
        讀取機器Dynamic或Static mode
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        string, Dynamic => Dynamic mode, Static => Static mode
        """ 
        s=self.sendQuery(f':MODE:DYNamic?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").lower()          
        if s in [str(self.Enum_Dynamic_Mode.Dynamic).lower(),'dynamic','dyn']:
            return self.Enum_Dynamic_Mode.Dynamic            
        elif s in [str(self.Enum_Dynamic_Mode.Static).lower(),'static','stat']:
            return self.Enum_Dynamic_Mode.Static      
        else:return None

    #設定機器CC Dynamic電流Level 1
    def setCC_Dynamic_Level_1(self,A,**kwargs):
        """
        設定機器CC Dynamic電流Level 1
        參數:
        R:float => 電流A ,MAX => 最大可設定電流, MIN => 最小可設定電流
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(A).upper().strip()        
        self.sendCommand(f':CURRent:L1 {s}',**kwargs)   

    #讀取機器CC Dynamic電流Level 1
    def getCC_Dynamic_Level_1(self,**kwargs):
        """
        讀取機器CC Dynamic電流Level 1
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電流A
        """ 
        s=self.sendQuery(f':CURRent:L1?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CC Dynamic電流Level 2
    def setCC_Dynamic_Level_2(self,A,**kwargs):
        """
        設定機器CC Dynamic電流Level 2
        參數:
        R:float => 電流A ,MAX => 最大可設定電流, MIN => 最小可設定電流
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(A).upper().strip()        
        self.sendCommand(f':CURRent:L2 {s}',**kwargs)   

    #讀取機器CC Dynamic電流Level 2
    def getCC_Dynamic_Level_2(self,**kwargs):
        """
        讀取機器CC Dynamic電流Level 1
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電流A
        """ 
        s=self.sendQuery(f':CURRent:L2?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CC Dynamic Slew Rate (Rise)
    def setCC_Dynamic_Slew_Rise(self,mA_us,**kwargs):
        """
        設定機器CC Dynamic Slew Rate (Rise)
        參數:
        mA_us:float => Slew Rate (mA/us) ,MAX => 最大可設定Slew Rate, MIN => 最小可設定Slew Rate
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(mA_us).upper().strip()        
        self.sendCommand(f':CURRent:RISE {s}',**kwargs)   

    #讀取機器CC Dynamic Slew Rate (Rise)
    def getCC_Dynamic_Slew_Rise(self,**kwargs):
        """
        讀取機器CC Dynamic Slew Rate (Rise)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, Slew Rate (mA/us)
        """ 
        s=self.sendQuery(f':CURRent:RISE?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CC Dynamic Slew Rate (Fall)
    def setCC_Dynamic_Slew_Fall(self,mA_us,**kwargs):
        """
        設定機器CC Dynamic Slew Rate (Fall)
        參數:
        mA_us:float => Slew Rate (mA/us) ,MAX => 最大可設定Slew Rate, MIN => 最小可設定Slew Rate
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(mA_us).upper().strip()        
        self.sendCommand(f':CURRent:FALL {s}',**kwargs)   

    #讀取機器CC Dynamic Slew Rate (Fall)
    def getCC_Dynamic_Slew_Fall(self,**kwargs):
        """
        讀取機器CC Dynamic Slew Rate (Fall)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, Slew Rate (mA/us)
        """ 
        s=self.sendQuery(f':CURRent:FALL?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CC Dynamic Timer (T1)
    def setCC_Dynamic_Timer_T1(self,sec,**kwargs):
        """
        設定機器CC Dynamic Timer (T1)
        參數:
        sec:float => timer (seconds) ,MAX => 最大可設定Slew Rate, MIN => 最小可設定Slew Rate
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(sec).upper().strip()        
        self.sendCommand(f':CURRent:T1 {s}',**kwargs)   

    #讀取機器CC Dynamic Slew Rate (Rise)
    def getCC_Dynamic_Timer_T1(self,**kwargs):
        """
        讀取機器CC Dynamic Timer (T1)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, timer (seconds)
        """ 
        s=self.sendQuery(f':CURRent:T1?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #設定機器CC Dynamic Timer (T2)
    def setCC_Dynamic_Timer_T2(self,sec,**kwargs):
        """
        設定機器CC Dynamic Timer (T2)
        參數:
        sec:float => timer (seconds) ,MAX => 最大可設定Slew Rate, MIN => 最小可設定Slew Rate
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(sec).upper().strip()        
        self.sendCommand(f':CURRent:T2 {s}',**kwargs)   

    #讀取機器CC Dynamic Slew Rate (Rise)
    def getCC_Dynamic_Timer_T2(self,**kwargs):
        """
        讀取機器CC Dynamic Timer (T2)
        參數:        
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, timer (seconds)
        """ 
        s=self.sendQuery(f':CURRent:T2?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()            
        if isnumber(s):return float(s)
        else:return None

    #取得電流量測值
    def getMeasure_Current(self,**kwargs):
        """
        取得電流量測值
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電流A
        """        
        s=self.sendQuery(f':MEASure:CURRent?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
        if isnumber(s):return float(s)
        else:return None
    
    #取得電壓量測值
    def getMeasure_Voltage(self,**kwargs):
        """
        取得電壓量測值
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 電壓V
        """        
        s=self.sendQuery(f':MEASure:VOLTage?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
        if isnumber(s):return float(s)
        else:return None
    
    #取得Load On持續時間
    def getMeasure_Elapsed_Time(self,**kwargs):
        """
        取得Load On持續時間
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, elapsed time (seconds)
        """        
        s=self.sendQuery(f':MEASure:ETIMe?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
        if isnumber(s):return float(s)
        else:return None

    #取得功率量測值
    def getMeasure_Power(self,**kwargs):
        """
        取得功率量測值
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        float, 功率W
        """        
        s=self.sendQuery(f':MEASure:POWer?',**kwargs)
        s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
        if isnumber(s):return float(s)
        else:return None
    def INP_ONOFF(self):
        self.sendCommand(":INP ON")
    


#**********  PSW物件  **********
class DC_PS(Instrument):
 
        #使用枚舉來定義常數機器常用常數
        class Enum_Control_Mode():
            Local_Control=0         #0 | NONE Local (PanelPanel) control
            Ext_V=1                 #1 | VOLTage External voltage control
            Ext_R_Rising=2          #2 | RRISing External resistance control; 10kΩ or 5kΩ = Io max max*, 0kΩ = Io
            Ext_R_Falling=3         #3 | RFALling External resistance control; 10kΩor 5kΩ = Io min*, 0kΩ = Io max.
            Ext_V_Isolation=4       #4 | VISolation External voltage control(isolated)    
        
        class Enum_Control_Range():
            Low=0                   #LOW | 0 5V [5kΩ]
            High=1                  #HIGH| 1 10 V [ 10 kΩ]
        #列舉輸出模式
        class Enum_Output_Mode():
            CV='CV'                     #CV Mode
            CC='CC'                     #CC Mode
            OFF='OFF'                   #Output Off
        #列舉輸出狀態
        class Enum_Output_State():
            OFF=0                   #Output Off
            ON=1                    #Output ON
            
        
        def __init__(self,**kwargs):
            super().__init__(**kwargs)
            self.instrPassWord=''       #校證模式密碼
                    # 各廠牌的密碼列表
            self.dictBrandPassword = {"GW": "abc123", "PRIST": "abc123",
                                "NF": "8508", "KG": "8508", "GW-Acbel": "abc123",
                                "Keithley-2260A": "2260A", "Keithley-2260B": "2260B", "RS ISO-TECH": "abc123",
                                "Chiyoda": "8508", "TEXIO": "abc123", "RS PRO": "abc123", "GW-PRP": "abc123", "Takasago": "abc123"}
            self.Volt_Max=None          #機器最大電壓, PSU可設定電壓 0~105% of the rated output voltage in volts.
            self.Volt_Rated=None        #機器額定電壓
            self.Volt_Min=None          #機器最小電壓
            self.Current_Max=None       #機器最大電流, PSU可設定電流 0~105% of the rated current output level.
            self.Current_Rated=None     #機器額定電流
            self.Current_Min=None       #機器最小電流
        
        # #解析參數中前置,後置delay時間
        # def get_delay_args(self,args):
        #     delay_before=0
        #     delay_after=0
        #     for key, value in args:   #讀取動態參數
        #         print(f"{key}: {value}")
        #         if key.lower() in ["delay_before","wait_before"]:
        #             delay_before=value
        #         elif key.lower() in ["delay","delay_after","wait_after"]:
        #             delay_after=value
        #     return delay_before,delay_after
        
        #設定輸出狀態
        def setOutput_State(self,OnOff,**kwargs):
            """
            設定機器輸出ON或OFF
            參數:
            OnOff: 1, on => 輸出ON, 0, off => 輸出OFF
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            try: 
                delay_before,delay_after=self.get_delay_args(kwargs.items())   
                self.delay(delay_before)        #先等待指定時間,再執行
                s=str(OnOff).lower().strip()
                if s in [str(self.Enum_Output_State.ON).lower(), 'on','1']:
                    strcmd=f':OUTP ON'
                elif s in [str(self.Enum_Output_State.OFF).lower(), 'off','0']:
                    strcmd=f':output OFF'
                self.sendCommand(strcmd)
                self.delay(delay_after)         #等待指定時間, 再結束
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e
        #讀輸出狀態
        def getOutput_State(self,**kwargs):
            """
            讀取機器輸出狀態
            參數:        
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            integer, 1 => 輸出ON, 0 => 輸出OFF
            """
            try:            
                delay_before,delay_after=self.get_delay_args(kwargs.items())   
                self.delay(delay_before)        #先等待指定時間,再執行
                s=self.sendQuery(':output?')
                s=s.lower().strip().replace("\n","").replace("\r","").replace(" ","")
                self.delay(delay_after)         #等待指定時間, 再結束
                if s in ['true','1','on']:
                    return self.Enum_Output_State.ON
                else:
                    return self.Enum_Output_State.OFF
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e
            
        #取得機器的電壓電流的最大,最小可設定值
        def get_Volt_Current_Max_Min(self,**kwargs):
            """
            取得機器的電壓電流的最大,最小可設定值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            一次返回4個值, 電壓最大值, 電壓最小值, 電流最大值, 電流最小值
            同時會更新下列 
            self.Volt_Max,self.Volt_Min,self.Current_Max,self.Current_Min
            """
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items())   
                self.delay(delay_before)        #先等待指定時間,再執行
                self.setOutput_State(0,delay=0.1)       #闗閉輸出,delay指定秒,再執行下一步步
                v=self.getVoltage()     #讀取原本的電壓設定
                a=self.getCurrent()     #讀取原本的電流設定            
                self.setCurrent(0.5,delay=0.1)            
                self.setVoltage('MAX',delay=0.1)
                self.Volt_Max=self.getVoltage()
                self.setVoltage('MIN',delay=0.1)
                self.Volt_Min=self.getVoltage()
                self.setVoltage(1,delay=0.1)
                self.setCurrent('MAX',delay=0.1)
                self.Current_Max=self.getCurrent()
                self.setCurrent('MIN',delay=0.1)
                self.Current_Min=self.getCurrent()
                self.delay(delay_after)         #等待指定時間, 再結束
                return self.Volt_Max,self.Volt_Min,self.Current_Max,self.Current_Min
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e


        #取得機器的電壓電流的額定值
        def get_Volt_Current_Rated(self,**kwargs):
            """
            取得機器的電壓電流的最大,最小可設定值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            一次返回2個值, 電壓額定值, 電流額定值
            同時會更新下列 
            self.Volt_Rated, self.Current_Rated
            """
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items())   
                self.delay(delay_before)        #先等待指定時間,再執行
                self.setOutput_State(0,delay=0.1)       #闗閉輸出,delay指定秒,再執行下一步步
                v=self.getOVP()     #讀取原本的電壓設定
                a=self.getOCP()     #讀取原本的電流設定            
                # self.setCurrent(0.5,delay=0.1)            
                self.setOVP('MAX',delay=0.1)
                self.Volt_Rated=round(self.getOVP()/1.1,2)      #取到小數後二位
                
                self.setVoltage(1,delay=0.1)
                self.setOCP('MAX',delay=0.1)
                self.Current_Rated=round(self.getOCP()/1.1,2)   #取到小數後二位
                self.setOVP(v,delay=0.1)
                self.setOCP(a,delay=0.1)
                
                self.delay(delay_after)         #等待指定時間, 再結束
                return self.Volt_Rated,self.Current_Rated
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e
            
        #同時設定輸出電壓電流
        def setV_A(self,V,A,**kwargs):
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items())   
                self.delay(delay_before)        #先等待指定時間,再執行
                s=str(V).lower().strip()
                y=str(A).lower().strip()
                strcmd=f':APPL {s},{y}'
                self.sendCommand(strcmd)
                self.delay(delay_after)         #等待指定時間, 再結束
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e

        #設定輸出電壓
        def setVoltage(self,V,**kwargs):
            """
            設定輸出電壓
            參數:
            V:數值 => 電壓V ,MAX => 最大可設定電壓(額定電壓105%), MIN => 最小可設定電壓(額定電壓0%)
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items())   
                self.delay(delay_before)        #先等待指定時間,再執行
                s=str(V).lower().strip()
                strcmd=f':VOLTage {s}'
                self.sendCommand(strcmd)
                self.delay(delay_after)         #等待指定時間, 再結束
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e
        
        #讀輸出電壓設定值
        def getVoltage(self,**kwargs):
            """
            取得電壓設定值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 電壓V
            """
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items()) 
                self.delay(delay_before)        #先等待指定時間,再執行
                strcmd=':VOLTage?'
                s=self.sendQuery(strcmd)
                s=s.strip().replace("\n","").replace('\r','').replace(" ","")
                self.delay(delay_after)         #等待指定時間, 再結束
                if isnumber(s):
                    return float(s)
                else:
                    return None
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e    
        
        #設定輸出電流
        def setCurrent(self,A,**kwargs):
            """
            設定輸出電流
            參數:
            A:數值 => 電流A ,MAX => 最大可設定電流(額定電流105%), MIN => 最小可設定電流(額定電流0%)
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items()) 
                self.delay(delay_before)
                s=str(A).lower().strip()
                strcmd=f':CURRent {A}'
                self.sendCommand(strcmd)
                self.delay(delay_after)
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e
        #讀輸出電壓設定值
        def getCurrent(self,**kwargs):
            """
            取得電流設定值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 電流A
            """
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items()) 
                self.delay(delay_before)
                strcmd=':CURRent?'
                s=self.sendQuery(strcmd)
                s=s.strip().replace("\n","").replace('\r','').replace(" ","")
                self.delay(delay_after)
                if isnumber(s):
                    return float(s)
                else:
                    return None
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} visa.errors.VisaIOError: {e.description}\n'
                raise e       
            except Exception as e:
                self.ErrorMessages+=f'{inspect.currentframe().f_code.co_name} other exception: {str(e)}\n'
                raise e
        #設定OVP電壓,Maximum: Vrated * 1.1 
        def setOVP(self, V,**kwargs):
            """
            設定OVP電壓
            參數:
            V:數值 => 電壓V ,MAX => 最大可設定OVP電壓(額定電壓110%), MIN => 最小可設定OVP電壓(額定電壓*0.1>5V, 為5V, 否則為額定電壓*0.1V)
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            self.sendCommand(f':VOLT:PROT {V}',**kwargs)
        #讀取OVP設定值
        def getOVP(self,**kwargs):
            """
            取得OVP電壓設定值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 電壓V
            """
            s=self.sendQuery(f':VOLT:PROT?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")
            if isnumber(s): return float(s)                
            else: return None

        #設定OCP電流,Maximum:   Irated * 1.1 
        def setOCP(self, A,**kwargs):
            """
            設定OCP電流
            參數:
            A:數值 => 電流A ,MAX => 最大可設定OCP電流(額定電流110%), MIN => 最小可設定OCP電流(額定電流*0.1>5A, 為5A, 否則為額定電流*0.1A)
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            self.sendCommand(f':CURR:PROT {A}',**kwargs)
        
        #讀取OCP設定值
        def getOCP(self,**kwargs):
            """
            取得OCP電流設定值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 電流A
            """
            s=self.sendQuery(f':CURR:PROT?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")
            if isnumber(s): return float(s)                
            else: return None
        
        #設定CC control mode (local control( panel), external voltage control, external resistance control).
        def setCC_Control_Mode(self, Control_Mode,**kwargs):
            """
            設定CC控制模式, 本機面板操作, 外部電壓,電阻控制
            參數:
            Control_Mode:
                0, none         =>  本機面板操作
                1, voltage      =>  外部電壓
                2, rrising      =>  外部電阻(rising)
                3, rfalling     =>  外部電阻(falling)
                4, visolation   =>  外部電壓(isolation)
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            if not isnumber(Control_Mode):
                Control_Mode=str(Control_Mode).strip().lower()
            par=''
            if Control_Mode in [str(self.Enum_Control_Mode.Local_Control).lower, '0', '','none']:
                par=0
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_V).lower(), '1', 'volt','voltage']:
                par=1
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_R_Rising).lower(), '2','rris','rrising']:
                par=2
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_R_Falling).lower(), '3','rfal','rfalling']:
                par=3 
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_V_Isolation).lower(), '4','vis','visolation']:
                par=4
            self.sendCommand(f':SYSTem:CONFigure:CURRent:CONTrol {par}',**kwargs)  

        #讀取CC control mode (local control( panel), external voltage control, external resistance control).
        def getCC_Control_Model(self,**kwargs):
            """
            取得CC控制模式, 本機面板操作, 外部電壓,電阻控制
            參數:        
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            integer
                0   =>  本機面板操作
                1   =>  外部電壓
                2   =>  外部電阻(rising)
                3   =>  外部電阻(falling)
                4   =>  外部電壓(isolation)
            """
            s=self.sendQuery(f':SYSTem:CONFigure:CURRent:CONTrol?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")
            if s in ['0']:
                return self.Enum_Control_Mode.Local_Control
            elif s in ['1']:
                return self.Enum_Control_Mode.Ext_V
            elif s in ['2']:
                return self.Enum_Control_Mode.Ext_R_Rising
            elif s in ['3']:
                return self.Enum_Control_Mode.Ext_R_Falling
            elif s in ['4']:
                return self.Enum_Control_Mode.Ext_V_Isolation
            else:
                return None
            
        #設定CV control mode (local control( panel), external voltage control, external resistance control).
        def setCV_Control_Mode(self, Control_Mode,**kwargs):
            """
            設定CV控制模式, 本機面板操作, 外部電壓,電阻控制
            參數:
            Control_Mode:
                0, none         =>  本機面板操作
                1, voltage      =>  外部電壓
                2, rrising      =>  外部電阻(rising)
                3, rfalling     =>  外部電阻(falling)
                4, visolation   =>  外部電壓(isolation)
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            if not isnumber(Control_Mode):
                Control_Mode=str(Control_Mode).strip().lower()
            par=''
            if Control_Mode in [str(self.Enum_Control_Mode.Local_Control).lower(), '0', '','none']:
                par=0
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_V).lower(), '1', 'volt','voltage']:
                par=1
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_R_Rising).lower(), '2','rris','rrising']:
                par=2
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_R_Falling).lower(), '3','rfal','rfalling']:
                par=3 
            elif Control_Mode in [str(self.Enum_Control_Mode.Ext_V_Isolation).lower(), '4','vis','visolation']:
                par=4
            self.sendCommand(f':SYSTem:CONFigure:VOLTage:CONTrol {par}',**kwargs)  

        #讀取CC control mode (local control( panel), external voltage control, external resistance control).
        def getCV_Control_Model(self,**kwargs):
            """
            取得CV控制模式, 本機面板操作, 外部電壓,電阻控制
            參數:        
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            integer,
                0   =>  本機面板操作
                1   =>  外部電壓
                2   =>  外部電阻(rising)
                3   =>  外部電阻(falling)
                4   =>  外部電壓(isolation)
            """
            s=self.sendQuery(f':SYSTem:CONFigure:VOLTage:CONTrol?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")
            if s in ['0']:
                return self.Enum_Control_Mode.Local_Control
            elif s in ['1']:
                return self.Enum_Control_Mode.Ext_V
            elif s in ['2']:
                return self.Enum_Control_Mode.Ext_R_Rising
            elif s in ['3']:
                return self.Enum_Control_Mode.Ext_R_Falling
            elif s in ['4']:
                return self.Enum_Control_Mode.Ext_V_Isolation
            else:
                return None

        #設定CV control mode (local control( panel), external voltage control, external resistance control).
        def setControl_Range(self, Control_Range,**kwargs):
            """
            設定CV,CC控制模式的Range
            參數:
            Control_Range:
                0, low       =>  5V [5kΩ]
                1, high      =>  10V [ 10 kΩ]            
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:無
            """
            if not isnumber(Control_Range):
                Control_Mode=str(Control_Mode).strip().lower()
            par=''
            if Control_Mode in [str(self.Enum_Control_Range.Low).lower(), '0', 'low']:
                par=0
            elif Control_Mode in [str(self.Enum_Control_Range.High).lower(), '1', 'high']:
                par=1
            self.sendCommand(f':SYSTem:CONFigure:CONTrol:RANG {par}',**kwargs)  

        #讀取CC control mode (local control( panel), external voltage control, external resistance control).
        def getControl_Range(self,**kwargs):
            """
            取得CV,CC控制模式的Range
            參數:                    
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            integer,
                0   =>  5V [5kΩ]
                1   =>  10V [ 10 kΩ]
            """
            s=self.sendQuery(f':SYSTem:CONFigure:CONTrol:RANG?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")
            if s in [0,'0']:
                return self.Enum_Control_Range.Low
            elif s in [1,'1']:
                return self.Enum_Control_Range.High
            else:
                return None
        #讀取機器的輸出模式
        def getOutput_Mode(self,**kwargs):
            """
            取得機器的輸出模式
            參數:                    
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            string, CV=> CV Mode, CC=> CC Mode, OFF=>Output Off
            """
            s=self.sendQuery(f':SOUR:MODE?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","").upper()
            return s        #傳回值應為CV, CC, OFF,其中一個
        
        #清除機器的OVP,OCP,OTP保護
        def clear_Output_Protect(self, **kwargs):
            """
            清除機器的OVP,OCP,OTP保護
            參數:                    
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            """
            self.sendCommand(f':OUTPut:PROTection:CLEar',**kwargs)

        #讀輸出電壓量測值
        def getMeasure_Voltage(self,**kwargs):
            """
            取得電壓量測值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 電壓V
            """        
            s=self.sendQuery(f':MEASure:VOLTage?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
            if isnumber(s):return float(s)
            else:return None
        
        #取得電流量測值
        def getMeasure_Current(self,**kwargs):
            """
            取得電流量測值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 電流A
            """        
            s=self.sendQuery(f':MEASure:CURRent?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
            if isnumber(s):return float(s)
            else:return None
        
        #讀輸出電壓電流量測值
        def getMeasure_Voltage_Current(self,**kwargs):
            """
            取得電壓及電流量測值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float,float 電壓V,電流A
            """        
            s=self.sendQuery(f':MEASure:ALL?',**kwargs)
            ss=s.strip().replace("\n","").replace('\r','').replace(" ","").split(',')
            sv=None
            sa=None
            if len(ss)==2:            
                if isnumber(ss[0]):sv= float(ss[0])
                if isnumber(ss[1]):sa= float(ss[1])
            return sv,sa
                
        #取得功率量測值
        def getMeasure_Power(self,**kwargs):
            """
            取得功率量測值
            參數:
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            返回值:
            float, 功率W
            """        
            s=self.sendQuery(f':MEASure:POWer?',**kwargs)
            s=s.strip().replace("\n","").replace('\r','').replace(" ","")            
            if isnumber(s):return float(s)
            else:return None
            

        
        #寫入序號到機器
        def writeSerNo(self,SerNo="",**kwargs):
            SubName="writeSerNo"        
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items()) 
                self.delay(delay_before)
                strcmd = "ZCALib:APASS \""+ self.instrPassWord +"\""
                self.sendCommand(strcmd)  # 送出Password
                strcmd = "ZCALib:PROG \""+ SerNo +"\""  # 寫入序號
                self.sendCommand(strcmd)
                self.delay(delay_after)
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
                raise e       
            except Exception as e:
                self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
                raise e
        
        #寫入MAC到機器
        def writeMAC(self, MAC="",**kwargs):
            SubName="writeMAC"
            strcmd = ""
            strTemp = MAC
            strTemp = strTemp.strip()
            strMac = ""
            try:
                delay_before,delay_after=self.get_delay_args(kwargs.items()) 
                self.delay(delay_before)
                if strTemp != "":  # 寫入MAC                            
                    strTemp = strTemp.upper()            
                    if len(strTemp) == 12:
                        i = 0
                        while i < 12:
                            strMac += strTemp[i]
                            strMac += strTemp[i+1]
                            strMac += "-"
                            i += 2
                        i = len(strMac)
                        strMac = strMac[0:i-1]
                        strcmd = "ZCALib:APASS \""+ self.instrPassWord +"\""
                        self.sendCommand(strcmd)  # 送出Password
                        strcmd = "ZCALib:MAC \""+strMac+"\""
                        self.sendCommand(strcmd)  # 寫入MAC                                
                    else:                    
                        raise ValueError(f"錯誤: MAC='{MAC}', MAC字串長度不對。")
                self.delay(delay_after)
            except visa.errors.VisaIOError as e:
                self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
                raise e
            except Exception as e:
                self.ErrorMessages+=SubName+" other exception: " + str(e) +"\n"
                raise e
                    
        #重開機
        def Reboot(self,**kwargs):
            """
            機器重啟動
            參數:                    
            **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
            """
            self.sendCommand(":SYSTem:REBoot",**kwargs)





#**********  MEGA2560物件  **********
class MEGA2560_GPT(Instrument):
    def get_IDN(self):
        id=self.sendQuery('*idn?')
        return id       
    
    def get_LastCMD(self):
        s=self.sendQuery('lastcmd?')
        return s
    

    
    

    
    
    