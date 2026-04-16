import os, sys
import shutil
import win32gui
import win32con
from time import sleep
from win32api import ShellExecute

class frmUpdate(object):    
    def __init__(self):
        super(frmUpdate,self).__init__()
        self.isClosing=False

    def onQApplicationStarted(self):
        try:
            # 更新源目錄
            update_dir = "\\\\172.16.23.129\\upload$\\TPE\\TestPro_Upgrade\\GPT98_99ATE\\"
            # 目標目錄
            current_dir = os.path.dirname(__file__)
            current_file=os.path.basename(__file__)
            file_exe=current_file.lower().replace('update_','').replace('.py','.exe')
            print('Update_GPT98_99ATE準備下載新程式...')
            print(f'目前程式資料夾:{current_dir}')
            sleep(2)
            self.closeMainPro()
        except Exception as e:
            print(str(e))
        sleep(3)
        # 遍歷更新源目錄中的所有檔案
        for root, dirs, files in os.walk(update_dir):
            for file in files:
                try:            
                    # 獲取源檔案的完整路徑
                    update_file_path = os.path.join(root, file)
                    if update_file_path.lower().find(current_file)>=0:
                        continue
                    # 計算源檔案相對於更新源目錄的相對路徑
                    relative_path = os.path.relpath(update_file_path, update_dir)
                    # 獲取目標檔案的完整路徑
                    current_file_path = os.path.join(current_dir, relative_path)
                    # 如果目標檔案不存在，或者源檔案的修改時間較新，則複製檔案
                    if not os.path.exists(current_file_path) or os.path.getmtime(update_file_path) > os.path.getmtime(current_file_path):
                        # 創建目標檔案所在的目錄
                        os.makedirs(os.path.dirname(current_file_path), exist_ok=True)
                        # 複製檔案
                        print(f'下載{update_file_path}')
                        # shutil.copy2(update_file_path, current_file_path)
                        shutil.copy(update_file_path, current_dir)
                        # QCoreApplication.processEvents()
                except Exception as e:
                    print(str(e))
        sleep(2)
        try:
            # 啟動新版本的程式
            # subprocess.Popen(f"{current_dir}\\OfficeKanban.exe")        
            print(f'啟動{current_dir}\\{file_exe}...')        
            ShellExecute(0, 'open', f"{current_dir}\\{file_exe}", '', '', 1) #打開主程式
            print(f'關閉更新程式本身...')
        except Exception as e:
            print(str(e))
        sleep(5)
        sys.exit(0)

    #關閉Update程式
    def closeMainPro(self):
        try:
            current_file_main=os.path.basename(__file__)
            current_file_main=current_file_main.lower()
            current_file_main=current_file_main.replace('.exe','')
            hWndList = []
            win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)  #列出所有視窗    
            i=0
            for hwnd in hWndList:       #關閉之前開啟的smart evision
                try:
                    # if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                        clsname = win32gui.GetClassName(hwnd)
                        title = win32gui.GetWindowText(hwnd)
                        if ((str(title).lower().find('GPT-98_99_1000')>=0 ) and str(title).lower().find(current_file_main)<0 ):
                            print(f"關閉之前開啟的主程式. \nhwnd:{hwnd}, class:{clsname}, title:{title}")
                            sleep(0.5) 
                            win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
                            sleep(1)
                except Exception as e:                
                    print(str(e))
                    continue 
        except Exception as e:                
            print(str(e))
    
    
    

if __name__ == '__main__':
    try:        
        window = frmUpdate()
        window.onQApplicationStarted()
    except Exception as e:        
        print(f'Error:{str(e)}')
