@echo off
rem 檢查是否存在 *.py 檔案
if exist *.py (  
  rem 存在 *.py 檔案，則結束批次檔
  echo 開發環境資料夾(有py檔), 不續繼進行更新作業. 
  exit
)

@echo off
echo 3秒後, 開始進行更新...
timeout /t 3 > NUL

echo 改變資料夾為本程式資料夾
CD /D "%~dp0"

echo 刪除當前資料夾中的子資料夾
for /d %%d in ("*") do (  
  rmdir "%%d" /S /Q
)

echo 刪除目前資料夾中不是Upgrade_GPT-ATE.bat的檔案...
for %%f in (".\*") do (  
  if not "%%f" == ".\Upgrade_GPT-ATE.bat" (
    del "%%f"
  )
)

echo 2秒後, 開始下載檔案...
timeout /t 2 > NUL

echo 下載檔案...
xcopy \\172.16.23.129\upload$\TPE\TestPro_Upgrade\GPT98_99ATE\*.* /D /E /Y

echo 3秒後, 開啟GPT98_99ATE.exe...
timeout /t 3 > NUL

echo 開啟 GPT98_99ATE.exe...
start .\GPT98_99ATE.exe

echo 2秒後,關閉本程式...
timeout /t 2 > NUL

end:

exit