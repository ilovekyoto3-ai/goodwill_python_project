import pyvisa as visa
from pyvisa import constants
from time import *
from fnmatch import fnmatch
from PyQt5 import QtWidgets
# import pyvisa
import win32com.client
import inspect

#from MyModules.Base_Instr import Instrument
#from MyModules.MyFunctions import isnumber
from enum import Enum   

#**********  DC Load GW_PEL-3000物件  **********
class LOAD_PEL_3K():
    #列舉輸入狀態
    class Enum_Input_State(Enum):
        OFF=0                   #input Off (load off)
        ON=1                    #input on (load on)
    
    #列舉輸入短路狀態
    class Enum_Input_Short_State(Enum):
        OFF=0                   #input short Off
        ON=1                    #input short on
    
    #列舉操作模式CC,CR,CV,CP...
    class Enum_Oper_Mode(Enum):
        CC='CC' #CC Mode
        CR='CR' #CR Mode
        CV='CV' #CV Mode
        CP='CP' #CP Mode
        CCCV='CCCV' #CC+CV Mode
        CRCV='CRCV' #CR+CV Mode
        CPCV='CPCV' #CP+CV Mode

    #列舉操作模式CC,CR,CV,CP...
    class Enum_Dynamic_Mode(Enum):
        Dynamic='Dynamic'   #Dynamic Mode
        Static='Static'     #Static Mode        
    
    #列舉電流檔位
    class Enum_Current_Range(Enum):
        Low='LOW'       #Low range
        Middle='MIDDLE' #Middle range
        High='HIGH'     #High range
    
    #列舉電壓檔位
    class Enum_Voltage_Range(Enum):
        Low='LOW' #Low range
        High='HIGH' #High range
        

    def __init__(self):
        super().__init__()
    
    #設定機器輸入ON或OFF
    def setInput_State(self,OnOff,**kwargs):
        """
        設定機器輸入ON或OFF (Load ON/OFF)
        參數:
        OnOff: 1, on => 輸入ON, 0, off => 輸入OFF
        **kwargs: delay_before=前置等待時間(秒),  delay_after 或 delay=程式結束前等待時間(秒)
        返回值:無
        """
        s=str(OnOff).lower().strip()
        strcmd=None
        if s in [self.Enum_Input_State.ON, 'on','1']:
            strcmd=f':INPut ON'
        elif s in [self.Enum_Input_State.OFF, 'off','0']:
            strcmd=f':INPut OFF'
        if strcmd is not None:
            self.sendCommand(strcmd,**kwargs)   

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
    def setInput_State(self,OnOff,**kwargs):
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
    def getInput_State(self,**kwargs):
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
        elif s in [str(self.Enum_Current_Range.Middle).upper(), 'MIDDLE','MIDD']:
            s1=f'MIDDLE'
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
        elif s in [str(self.Enum_Current_Range.Middle).upper(),'MIDDLE','MIDD']:
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


