@echo off
chcp 65001
echo ====================================
echo   糖尿病品質管理 Dashboard 啟動程式
echo ====================================

REM 檢查 Python 是否已安裝
python --version > nul 2>&1
if errorlevel 1 (
    echo [錯誤] 未安裝 Python，請先安裝 Python 3.9 或更新版本
    echo 可從 https://www.python.org/downloads/ 下載
    pause
    exit /b
)

REM 安裝必要套件
echo [步驟 1] 檢查並安裝必要套件...
pip install -r requirements.txt
if errorlevel 1 (
    echo [錯誤] 安裝套件失敗
    pause
    exit /b
)

echo [步驟 2] 啟動應用程式...
echo 程式將在瀏覽器中開啟，請稍候...

REM 創建啟動器批次檔
echo [步驟 3] 創建桌面捷徑...
set "LAUNCHER=%~dp0launcher.vbs"
(
echo Set WShell = CreateObject^("WScript.Shell"^)
echo WShell.CurrentDirectory = "%~dp0"
echo WShell.Run "cmd /c streamlit run app.py", 0, False
) > "%LAUNCHER%"

REM 創建捷徑
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\糖尿病品質管理Dashboard.lnk" >> CreateShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateShortcut.vbs
echo oLink.TargetPath = "%LAUNCHER%" >> CreateShortcut.vbs
echo oLink.WorkingDirectory = "%~dp0" >> CreateShortcut.vbs
echo oLink.IconLocation = "%SystemRoot%\System32\SHELL32.dll,144" >> CreateShortcut.vbs
echo oLink.Save >> CreateShortcut.vbs
cscript //nologo CreateShortcut.vbs
del CreateShortcut.vbs

REM 啟動 Streamlit
start streamlit run app.py

echo [完成] 桌面捷徑已創建，您可以使用捷徑直接開啟應用程式
pause 