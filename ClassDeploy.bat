@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

if /I "%~1"=="build-agent" goto :build_agent
if /I "%~1"=="install-server" goto :install_server
if /I "%~1"=="remove-server" goto :remove_server
if /I "%~1"=="run-server" goto :run_server
if /I "%~1"=="install-agent" goto :install_agent
if /I "%~1"=="remove-agent" goto :remove_agent

call :detect_python
if errorlevel 1 goto :end

:menu
cls
echo ============================================================
echo                    ClassDeploy Control
echo ============================================================
echo 1. Build agent EXE
echo 2. Install server
echo 3. Remove server
echo 4. Run server now
echo 5. Install agent
echo 6. Remove agent
echo 0. Exit
echo.
set "CHOICE="
set /p "CHOICE=Select action: "
if "%CHOICE%"=="1" goto :build_agent
if "%CHOICE%"=="2" goto :install_server
if "%CHOICE%"=="3" goto :remove_server
if "%CHOICE%"=="4" goto :run_server
if "%CHOICE%"=="5" goto :install_agent
if "%CHOICE%"=="6" goto :remove_agent
if "%CHOICE%"=="0" goto :end
goto :menu

:detect_python
set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
    echo Python 3 was not found in PATH.
    pause
    exit /b 1
)
exit /b 0

:ensure_python
if defined PYTHON_CMD exit /b 0
call :detect_python
exit /b %errorlevel%

:require_admin
net session >nul 2>&1
if errorlevel 1 (
    echo Run this BAT as Administrator.
    pause
    exit /b 1
)
exit /b 0

:install_deps
if exist "%~dp0vendor\wheels\*.whl" (
    echo Installing dependencies from vendor\wheels ...
    %PYTHON_CMD% -m pip install --no-index --find-links "%~dp0vendor\wheels" -r "%~dp0requirements.txt"
) else (
    echo Installing dependencies from PyPI ...
    %PYTHON_CMD% -m pip install -r "%~dp0requirements.txt"
)
exit /b %errorlevel%

:build_agent
call :ensure_python
if errorlevel 1 goto :end
cls
if exist "%~dp0build" rmdir /s /q "%~dp0build"
if exist "%~dp0dist" rmdir /s /q "%~dp0dist"
if exist "%~dp0ClassDeployAgent.spec" del /f /q "%~dp0ClassDeployAgent.spec" >nul 2>nul
if exist "%~dp0ClassDeployAgent.exe" del /f /q "%~dp0ClassDeployAgent.exe" >nul 2>nul
call :install_deps
if errorlevel 1 (
    echo Failed to install dependencies.
    pause
    goto :menu
)
echo.
echo Building ClassDeployAgent.exe ...
%PYTHON_CMD% -m PyInstaller --noconfirm --clean --onefile --name ClassDeployAgent --noconsole --add-data "agent\overlay.py;agent" --hidden-import win32timezone --hidden-import pythoncom --hidden-import pywintypes --collect-submodules pywinauto --collect-submodules pycaw --collect-submodules comtypes --collect-submodules win32com agent\main.py
if errorlevel 1 (
    echo Build failed.
    pause
    goto :menu
)
copy /Y "%~dp0dist\ClassDeployAgent.exe" "%~dp0ClassDeployAgent.exe" >nul
echo.
echo Build completed: ClassDeployAgent.exe
pause
goto :menu

:install_server
call :ensure_python
if errorlevel 1 goto :end
call :require_admin
if errorlevel 1 goto :menu
cls
call :install_deps
if errorlevel 1 (
    echo Failed to install dependencies.
    pause
    goto :menu
)
set "SERVER_DATA=%ProgramData%\ClassDeploy\server"
if not exist "%SERVER_DATA%" mkdir "%SERVER_DATA%"
netsh advfirewall firewall delete rule name="ClassDeploy Server 8765" >nul 2>nul
netsh advfirewall firewall add rule name="ClassDeploy Server 8765" dir=in action=allow protocol=TCP localport=8765 >nul 2>nul
call :make_server_shortcut
echo.
echo Server install completed.
pause
goto :menu

:make_server_shortcut
set "VBS_FILE=%TEMP%\classdeploy_server_shortcut.vbs"
>"%VBS_FILE%" echo Set oWS = WScript.CreateObject("WScript.Shell")
>>"%VBS_FILE%" echo sDesktop = oWS.SpecialFolders("AllUsersDesktop")
>>"%VBS_FILE%" echo Set oLink = oWS.CreateShortcut(sDesktop ^& "\ClassDeploy Server.lnk")
>>"%VBS_FILE%" echo oLink.TargetPath = "%~dp0start_server.bat"
>>"%VBS_FILE%" echo oLink.WorkingDirectory = "%~dp0"
>>"%VBS_FILE%" echo oLink.IconLocation = "%SystemRoot%\System32\SHELL32.dll,21"
>>"%VBS_FILE%" echo oLink.Save
cscript //nologo "%VBS_FILE%" >nul 2>nul
del /f /q "%VBS_FILE%" >nul 2>nul
exit /b 0

:remove_server
call :require_admin
if errorlevel 1 goto :menu
cls
netsh advfirewall firewall delete rule name="ClassDeploy Server 8765" >nul 2>nul
if exist "%ProgramData%\ClassDeploy\server" rmdir /s /q "%ProgramData%\ClassDeploy\server"
del /f /q "%UserProfile%\Desktop\ClassDeploy Server.lnk" >nul 2>nul
del /f /q "%Public%\Desktop\ClassDeploy Server.lnk" >nul 2>nul
echo.
echo Server data removed.
pause
goto :menu

:ensure_agent_exe
if exist "%~dp0ClassDeployAgent.exe" (
    set "AGENT_EXE=%~dp0ClassDeployAgent.exe"
    exit /b 0
)
if exist "%~dp0dist\ClassDeployAgent.exe" (
    set "AGENT_EXE=%~dp0dist\ClassDeployAgent.exe"
    exit /b 0
)
echo ClassDeployAgent.exe was not found. Build it first.
exit /b 1

:install_agent
call :require_admin
if errorlevel 1 goto :menu
cls
set "SERVER_IP="
set /p "SERVER_IP=Enter server IP: "
if "%SERVER_IP%"=="" (
    echo Server IP is required.
    pause
    goto :menu
)
call :ensure_agent_exe
if errorlevel 1 (
    pause
    goto :menu
)
set "DATA_DIR=%ProgramData%\ClassDeploy"
set "INSTALL_DIR=%DATA_DIR%\Agent"
set "RUN_REG_DATA=\"%SystemRoot%\System32\wscript.exe\" \"%INSTALL_DIR%\StartClassDeployAgent.vbs\""

if not exist "%DATA_DIR%" mkdir "%DATA_DIR%"
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
if not exist "%DATA_DIR%\logs" mkdir "%DATA_DIR%\logs"
if not exist "%DATA_DIR%\temp" mkdir "%DATA_DIR%\temp"

copy /Y "%AGENT_EXE%" "%INSTALL_DIR%\ClassDeployAgent.exe" >nul || (
    echo Failed to copy agent EXE.
    pause
    goto :menu
)
> "%DATA_DIR%\server.txt" echo %SERVER_IP%

call :make_agent_watchdog_files
if errorlevel 1 (
    echo Failed to create StartClassDeployAgent files.
    pause
    goto :menu
)

icacls "%DATA_DIR%" /grant *S-1-5-32-545:(OI)(CI)M /T /C >nul 2>nul
icacls "%INSTALL_DIR%" /grant *S-1-5-32-545:(OI)(CI)RX /T /C >nul 2>nul

reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /f >nul 2>nul
reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /t REG_SZ /d "%RUN_REG_DATA%" /f >nul 2>nul

call :write_startup_vbs "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Startup\ClassDeploy Agent.vbs"
call :write_existing_user_startup_files
call :write_default_user_startup_file
call :register_current_user_run
call :register_existing_user_run_entries
call :register_default_user_run_entry
call :register_active_setup

if not exist "%INSTALL_DIR%\StartClassDeployAgent.vbs" (
    echo StartClassDeployAgent.vbs was not created: %INSTALL_DIR%\StartClassDeployAgent.vbs
    pause
    goto :menu
)

taskkill /F /IM ClassDeployAgent.exe >nul 2>nul
start "" "%SystemRoot%\System32\wscript.exe" "%INSTALL_DIR%\StartClassDeployAgent.vbs"

echo.
echo Agent installed as interactive autostart for all users.
echo Current server: %SERVER_IP%:8765
echo Data directory: %DATA_DIR%
pause
goto :menu

:make_agent_watchdog_files
> "%INSTALL_DIR%\ClassDeployAgent.watch" echo run
> "%INSTALL_DIR%\StartClassDeployAgent.cmd" echo @echo off
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo setlocal EnableExtensions
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo if not exist "%%~dp0ClassDeployAgent.watch" exit /b 0
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo :loop
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo if not exist "%%~dp0ClassDeployAgent.watch" exit /b 0
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo start "" /wait "%%~dp0ClassDeployAgent.exe"
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo timeout /t 3 /nobreak ^>nul
>>"%INSTALL_DIR%\StartClassDeployAgent.cmd" echo goto loop
> "%INSTALL_DIR%\StartClassDeployAgent.vbs" echo Set oWS = CreateObject("WScript.Shell")
>>"%INSTALL_DIR%\StartClassDeployAgent.vbs" echo scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
>>"%INSTALL_DIR%\StartClassDeployAgent.vbs" echo cmdPath = scriptDir ^& "\StartClassDeployAgent.cmd"
>>"%INSTALL_DIR%\StartClassDeployAgent.vbs" echo If CreateObject("Scripting.FileSystemObject").FileExists(cmdPath) Then oWS.Run Chr(34) ^& cmdPath ^& Chr(34), 0, False
> "%INSTALL_DIR%\ClassDeployUserInit.cmd" echo @echo off
>>"%INSTALL_DIR%\ClassDeployUserInit.cmd" echo reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /t REG_SZ /d "\"%%SystemRoot%%\System32\wscript.exe\" \"%%~dp0StartClassDeployAgent.vbs\"" /f ^>nul 2^>nul
>>"%INSTALL_DIR%\ClassDeployUserInit.cmd" echo if exist "%%~dp0StartClassDeployAgent.vbs" start "" /min "%%SystemRoot%%\System32\wscript.exe" "%%~dp0StartClassDeployAgent.vbs"
if not exist "%INSTALL_DIR%\StartClassDeployAgent.cmd" exit /b 1
if not exist "%INSTALL_DIR%\StartClassDeployAgent.vbs" exit /b 1
if not exist "%INSTALL_DIR%\ClassDeployUserInit.cmd" exit /b 1
exit /b 0

:write_startup_vbs
set "TARGET_VBS=%~1"
if "%TARGET_VBS%"=="" exit /b 1
for %%I in ("%TARGET_VBS%") do if not exist "%%~dpI" mkdir "%%~dpI"
> "%TARGET_VBS%" echo Set fso = CreateObject("Scripting.FileSystemObject")
>>"%TARGET_VBS%" echo scriptPath = "%INSTALL_DIR%\ClassDeployUserInit.cmd"
>>"%TARGET_VBS%" echo If fso.FileExists(scriptPath) Then
>>"%TARGET_VBS%" echo^  CreateObject("WScript.Shell").Run Chr(34) ^& scriptPath ^& Chr(34), 0, False
>>"%TARGET_VBS%" echo End If
exit /b 0

:write_existing_user_startup_files
for /d %%D in ("%SystemDrive%\Users\*") do (
    set "USERNAME_DIR=%%~nxD"
    if /I not "!USERNAME_DIR!"=="Default" if /I not "!USERNAME_DIR!"=="Default User" if /I not "!USERNAME_DIR!"=="Public" if /I not "!USERNAME_DIR!"=="All Users" (
        call :write_startup_vbs "%%~fD\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\ClassDeploy Agent.vbs"
    )
)
exit /b 0

:write_default_user_startup_file
if exist "%SystemDrive%\Users\Default" (
    call :write_startup_vbs "%SystemDrive%\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\ClassDeploy Agent.vbs"
)
exit /b 0

:register_current_user_run
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /f >nul 2>nul
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /t REG_SZ /d "%RUN_REG_DATA%" /f >nul 2>nul
exit /b 0

:register_loaded_user_hive_run
set "HIVE_ROOT=%~1"
if "%HIVE_ROOT%"=="" exit /b 1
reg delete "%HIVE_ROOT%\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /f >nul 2>nul
reg add "%HIVE_ROOT%\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /t REG_SZ /d "%RUN_REG_DATA%" /f >nul 2>nul
exit /b 0

:register_existing_user_run_entries
set "HIVE_INDEX=0"
for /d %%D in ("%SystemDrive%\Users\*") do (
    set "USERNAME_DIR=%%~nxD"
    if /I not "!USERNAME_DIR!"=="Default" if /I not "!USERNAME_DIR!"=="Default User" if /I not "!USERNAME_DIR!"=="Public" if /I not "!USERNAME_DIR!"=="All Users" (
        if exist "%%~fD\NTUSER.DAT" (
            set /a HIVE_INDEX+=1
            set "HIVE_NAME=HKU\ClassDeployTemp!HIVE_INDEX!"
            reg load "!HIVE_NAME!" "%%~fD\NTUSER.DAT" >nul 2>nul
            if not errorlevel 1 (
                call :register_loaded_user_hive_run "!HIVE_NAME!"
                reg unload "!HIVE_NAME!" >nul 2>nul
            )
        )
    )
)
exit /b 0

:register_default_user_run_entry
if exist "%SystemDrive%\Users\Default\NTUSER.DAT" (
    reg load "HKU\ClassDeployDefault" "%SystemDrive%\Users\Default\NTUSER.DAT" >nul 2>nul
    if not errorlevel 1 (
        call :register_loaded_user_hive_run "HKU\ClassDeployDefault"
        reg unload "HKU\ClassDeployDefault" >nul 2>nul
    )
)
exit /b 0

:register_active_setup
reg delete "HKLM\Software\Microsoft\Active Setup\Installed Components\ClassDeployAgent" /f >nul 2>nul
reg add "HKLM\Software\Microsoft\Active Setup\Installed Components\ClassDeployAgent" /v "Version" /t REG_SZ /d "1,0,2,0" /f >nul 2>nul
reg add "HKLM\Software\Microsoft\Active Setup\Installed Components\ClassDeployAgent" /v "IsInstalled" /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKLM\Software\Microsoft\Active Setup\Installed Components\ClassDeployAgent" /v "StubPath" /t REG_SZ /d "\"%INSTALL_DIR%\ClassDeployUserInit.cmd\"" /f >nul 2>nul
exit /b 0

:remove_agent
call :require_admin
if errorlevel 1 goto :menu
cls
reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /f >nul 2>nul
reg delete "HKLM\Software\Microsoft\Active Setup\Installed Components\ClassDeployAgent" /f >nul 2>nul
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ClassDeployAgent /f >nul 2>nul
del /f /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Startup\ClassDeploy Agent.vbs" >nul 2>nul
del /f /q "%SystemDrive%\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\ClassDeploy Agent.vbs" >nul 2>nul
for /d %%D in ("%SystemDrive%\Users\*") do del /f /q "%%~fD\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\ClassDeploy Agent.vbs" >nul 2>nul
taskkill /F /IM ClassDeployAgent.exe >nul 2>nul
if exist "%ProgramData%\ClassDeploy\Agent\ClassDeployAgent.watch" del /f /q "%ProgramData%\ClassDeploy\Agent\ClassDeployAgent.watch" >nul 2>nul
if exist "%ProgramData%\ClassDeploy\Agent" rmdir /s /q "%ProgramData%\ClassDeploy\Agent"
if exist "%ProgramData%\ClassDeploy\logs" rmdir /s /q "%ProgramData%\ClassDeploy\logs"
if exist "%ProgramData%\ClassDeploy\temp" rmdir /s /q "%ProgramData%\ClassDeploy\temp"
if exist "%ProgramData%\ClassDeploy\server.txt" del /f /q "%ProgramData%\ClassDeploy\server.txt" >nul 2>nul
echo.
echo Agent removed.
pause
goto :menu

:run_server
call :ensure_python
if errorlevel 1 goto :end
set "CLASS_DEPLOY_DATA_DIR=%ProgramData%\ClassDeploy\server"
if not exist "%CLASS_DEPLOY_DATA_DIR%" mkdir "%CLASS_DEPLOY_DATA_DIR%"
cls
echo Starting server...
%PYTHON_CMD% -m server.main
echo.
pause
goto :menu

:end
endlocal
exit /b 0
