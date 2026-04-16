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
from MyModules.MyFunctions import isnumber

# def isnumber(aString):
#     try:
#         float(aString)
#         return True
#     except:
#         return False

#**********  儀器底層物件  **********
class Instrument():
    
    def __init__(self,**kwargs):
        self.flag_stop=False
        self.VisaDescription=""
        self.Baudrate=9600
        self.ErrorMessages=""
        self.instrOpened=False
        self.instrIDN=""
        self.instrTimeouts_ms=500
        self.instrDelayTime_s=0.1
        self.instrPassWord="abc123"
        self.instrConnected=0
        self.IDN_Keyword=""
        self.LastAnswer=''
        self.Device_Alias=''   

        self.rm = visa.ResourceManager()
        self.instr=visa.resources.Resource(self.rm,"")
        for key, value in kwargs.items():
            if key.lower()=="device_alias":
                self.Device_Alias=value
    
    def delay(self,delaySec=0.0):
        """
        延遲等待一段時間
        參數:
        delaySec:float => 要等待的時間(seconds)        
        返回值:無
        """
        if delaySec<=0: return
        if self.flag_stop: return
        wt=0
        st=monotonic()
        while(wt<delaySec):
            wt=monotonic()-st
            QtWidgets.QApplication.processEvents()
            sleep(0.001)    
    
    #以關鍵字找機器的通訊埠, 找到傳回Port的Visa Description
    def findPort(self, keyword="", command="*idn?"):
        """
        以關鍵字(keyword)找機器的通訊埠
        參數:
        keyword:string => 要查找機器回傳值的關鍵字
        command:string => 要讓機器回傳值的命令
        返回值:
        string,找到Port的Visa Description
        """        
        #rm = visa.ResourceManager()
        SubName=inspect.currentframe().f_code.co_name
        foundPort="" 
        dev=""
        s=""
        if keyword=="":
            keyword=self.IDN_Keyword

        if keyword=="":
            raise ValueError("當self.IDN_Keyword為空字串時, keyword不能為空字串.")
        else:
            strDev_2find=keyword.lower()
        if self.VisaDescription!="":
            if self.checkPort(self.VisaDescription,strDev_2find,command):
                foundPort=self.VisaDescription
        if foundPort=="":
            tupDevices=self.rm.list_resources()
            for dev in tupDevices:
                if self.checkPort(dev,strDev_2find,command):
                    foundPort=dev
                    break       
        
        if fnmatch(self.VisaDescription,"ASRL*"):
            self.instr = visa.resources.SerialInstrument(self.rm, self.VisaDescription)  #設定resource為序列埠儀器
        elif fnmatch(self.VisaDescription,"GPIB*"):
            self.instr =visa.resources.GPIBInstrument(self.rm, self.VisaDescription)     #設定resource為GPIB儀器
        elif fnmatch(self.VisaDescription,"TCPIP*SOCKET"):
            self.instr =visa.resources.TCPIPSocket(self.rm, self.VisaDescription)       #設定resource為TCPIP SOCKET儀器
        elif fnmatch(self.VisaDescription,"TCPIP*"):
            self.instr =visa.resources.TCPIPInstrument(self.rm, self.VisaDescription)     #設定resource為TCPIP儀器
        elif fnmatch(self.VisaDescription,"USB*RAW"):
            self.instr =visa.resources.USBRaw(self.rm, self.VisaDescription)            #設定resource為USB RAW儀器    
        elif fnmatch(self.VisaDescription,"USB*"):
            self.instr =visa.resources.USBInstrument(self.rm, self.VisaDescription)     #設定resource為USBTMC儀器
        else:
            self.instr =visa.resources.MessageBasedResource(self.rm, self.VisaDescription)
            
        return foundPort

    def getErrMessage(self):
        s=self.ErrorMessages
        self.ErrorMessages=''
        return s

    #檢查port傳回值是否符合要查的關鍵字
    def checkPort(self, DevVisaDescript="", keyWord="", commandString='*idn?'):
        """
        檢查port傳回值是否符合要查的關鍵字
        參數:
        keyword:string => 要查找回傳值的關鍵字
        commandString:string => 要讓機器回傳值的命令
        返回值:
        boolean,找到 => True, 没找到 => False
        """
        SubName=inspect.currentframe().f_code.co_name
        checkOK=False
        if commandString is None:
            commandString='*idn?'
        elif commandString=='':
            commandString='*idn?'
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
            
            try:           
                if self.Baudrate!=9600:
                    inst.baud_rate=self.Baudrate
            except:pass

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
            self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            raise e                
        finally:            
            return checkOK
    

    #開啟通訊埠
    def openPort(self,**kwargs):
        """
        開啟通訊埠
        參數:
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        boolean, True => 成功開啟, False =>  開啟失敗,過程發生錯誤       
        """        
        SubName=inspect.currentframe().f_code.co_name
        is_open=False
        self.instrOpened=False
        s=self.VisaDescription        
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items()) 
            self.delay(delay_before)        #先等待指定時間,再執行
            self.instr=self.rm.open_resource(self.VisaDescription)
            if s.lower().find('asrl')>=0:
                self.instr.baud_rate=self.Baudrate
                self.instr.timeout=self.instrTimeouts_ms
            elif s.lower().find('tcpip')>=0:
                self.instr.timeout=1000 
                self.instr.set_visa_attribute(constants.VI_ATTR_TERMCHAR_EN,constants.VI_TRUE)               
            else:
                self.instr.timeout=self.instrTimeouts_ms
            self.instrOpened=True
            is_open=True
            self.delay(delay_after)         #等待指定時間, 再結束
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
            #raise e       
        except Exception as e:
            self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            #raise e
        return is_open
    #關閉通訊埠
    def closePort(self):
        """
        關閉通訊埠
        參數:無
        返回值:無              
        """
        SubName=inspect.currentframe().f_code.co_name        
        try:
            self.instr.close()
            self.instrOpened=False
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=SubName+" visa.errors.VisaIOError: " + e.description +"\n"
            raise e       
        except Exception as e:
            self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            raise e

    #解析參數中前置,後置delay時間
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
    def readString(self,**kwargs):
        """
        讀回機器傳回訊息
        參數: 
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:
        string, 機器回傳字串
        """
        SubName=inspect.currentframe().f_code.co_name
        strAns=""         
        try:
            delay_before,delay_after=self.get_delay_args(kwargs.items())   
            self.delay(delay_before)        #先等待指定時間,再執行
            strAns=self.instr.read()
            self.delay(delay_after)         #等待指定時間, 再結束
        except visa.errors.VisaIOError as e:
            self.ErrorMessages+=f'{SubName}, visa.errors.VisaIOError: {e.description}\n'
            #raise e       
        except Exception as e:
            self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            #raise e
        finally:
            return strAns
    #讀回IDN
    def get_IDN(self,retry_times=1):
        """
        讀回機器IDN訊息,並同時修改instrIDN
        參數: 
        無
        返回值:
        string, 機器回傳IDN字串
        """        
        SubName=inspect.currentframe().f_code.co_name
        m=0
        fd=False 
        while(m<=retry_times and fd==False):            
            id=''
            try:
                id=str(self.sendQuery('*idn?')).replace('\r','').replace('\n','').strip()
            except Exception as e:
                self.ErrorMessages+=f'{SubName}, other exception: {str(e)}\n'
            if id!='':
                fd=True 
                break                   
            # if m>1:
            #     print(f"{self.VisaDescription} get_IDN m=>{m}")                    
            m+=1
            if m<=retry_times:
                self.delay(self.instrDelayTime_s)
        self.instrIDN=id
        return id 

    def CLS(self):
        """
        對機器送*CLS命令
        參數:無
        返回值:無
        """        
        self.sendCommand('*CLS')

    def RST(self):
        """
        對機器送*RST命令
        參數:無
        返回值:無
        """        
        self.sendCommand('*RST')

    def getDevInfo(self):
        """
        讀取USB CDC裝置訊息
        """
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
    
    




    


    

    
    
    