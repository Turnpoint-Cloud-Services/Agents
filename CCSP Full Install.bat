@echo off
setlocal ENABLEDELAYEDEXPANSION
goto :NEXT
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
CCSP Client Installation Batch File
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    Version 2.0 3/31/2020 J Lear

    This script is supplied by Enghouse Interactive as is and is intended as reference for how to install CCSP Client Services from the command prompt.
    It is meant to be run by placing the file in a dedicated folder on a Windows PC, right clicking on the the file, and selecting run as administrator
    The script will use powershell to download all of the following required MSI files either to the same directory as the batch file or to the Windows %Temp% folder
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/ClientInstallationService/CCSPClientInstallationService.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPClientReportingService.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPClientCommunicator.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPClientTrayApp.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPClientUploadsService.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPScreenRecordingService.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPSIPService.msi
          https://TouchPointHOSTNAME/TouchPoint/ClientServices/CCSPTouchPointConnector.msi
          https://AdminHOSTNAME/webadministrator/administrator.msi
          https://AdminHOSTNAME/cosmodesigner/cosmodesigner.msi
          https://TouchPointHOSTNAME/touchpoint/clientservices/vcredist_x86.exe
          https://TouchPointHOSTNAME/touchpoint/clientservices/Encoder_en.exe

    set the TouchPointHOSTNAME variable below to the publicly accessible server host name of the TouchPoint application server.
    set the AdminHOSTNAME variable below to the publicly accessible server host name of the WebAdministrator and CosmoDesigner server.
    set UseWindwosTemp=False to retain the msi files in the same as the batch file to prevent subsequent downloads
    set UseWindwosTemp=True to run the install and cleanup the MSI downloads
    set EnableLogging=True to generate a log file for the MSI uninstalls and installs.  The log file path is displayed on the console output.

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:NEXT
    set TouchPointHOSTNAME=webcc.hellospoke.com
    set AdminHOSTNAME=webcc.hellospoke.com

    set UseWindowsTemp=False
    set EnableLogging=False

    set usingSSL=True
    set protocol=https
    set wcfHttpPort=8000
    set wcfHttpsPort=8001


:::: Start running Install.bat ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    cls
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo Running CCSP Client Installation ......................................................................................
    echo.

    set CD=%~d0%~p0
    for /f "tokens=1,2,3 delims=." %%a in ("%TouchPointHOSTNAME%") do set DOMAIN=%%b.%%c

    echo Running %~n0%~x0 from %CD%
    echo.
    echo CCSP Domain is %DOMAIN%
    echo CCSP TouchPoint URL is %protocol%://%TouchPointHOSTNAME%/TouchPoint
    echo.

    if "%UseWindowsTemp%" == "True" (
       set MSILocation=!Temp!
       set LOGDIR=!Temp!\CCSPClientInstallLog
    ) else (
       set MSILocation=!CD!
       set LOGDIR=!CD!CCSPClientInstallLog\!COMPUTERNAME!!USERNAME!
    )
    echo MSI location set to %MSILocation%


    if "%EnableLogging%" == "True" (
       echo MSI logging set to !LOGDIR!\install.log

       set MSILOGGING=/l*v+ "!LOGDIR!\install.log" 
       set VCRedistLog=/l "!LOGDIR!\vcredist_x86.log"
       set EncoderLog=-L*V:"!LOGDIR!\Encoder_en.log"

       if not exist "!LOGDIR!" md "!LOGDIR!"
       if exist "!LOGDIR!\install.log" del "!LOGDIR!\install.log" /Q
       if exist "!LOGDIR!\vcredist_x86.log" del "!LOGDIR!\vcredist_x86.log" /Q
       if exist "!LOGDIR!\Encoder_en.log" del "!LOGDIR!\Encoder_en.log" /Q
    ) else (
       echo Logging Dissabled
    )
    echo.
    echo.


:::: Downloading Files ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   echo Checking Client Installation MSI files ................................................................................
    if not "%UseWindowsTemp%" == "True" goto :NODEL
    if exist "%MSILocation%\CCSPClientInstallationService.msi" DEL "%MSILocation%\CCSPClientInstallationService.msi" /Q
    if exist "%MSILocation%\CCSPClientReportingService.msi"    DEL "%MSILocation%\CCSPClientReportingService.msi" /Q
    if exist "%MSILocation%\CCSPClientCommunicator.msi"        DEL "%MSILocation%\CCSPClientCommunicator.msi" /Q
    if exist "%MSILocation%\CCSPClientTrayApp.msi"             DEL "%MSILocation%\CCSPClientTrayApp.msi" /Q
    if exist "%MSILocation%\CCSPClientUploadsService.msi"      DEL "%MSILocation%\CCSPClientUploadsService.msi" /Q
    if exist "%MSILocation%\CCSPScreenRecordingService.msi"    DEL "%MSILocation%\CCSPScreenRecordingService.msi" /Q
    if exist "%MSILocation%\CCSPSIPService.msi"                DEL "%MSILocation%\CCSPSIPService.msi" /Q
    if exist "%MSILocation%\CCSPTouchPointConnector.msi"       DEL "%MSILocation%\CCSPTouchPointConnector.msi" /Q
    if exist "%MSILocation%\Encoder_en.exe"                    DEL "%MSILocation%\Encoder_en.exe" /Q
    if exist "%MSILocation%\vcredist_x86.exe"                  DEL "%MSILocation%\vcredist_x86.exe" /Q
    if exist "%MSILocation%\Administrator.msi"                 DEL "%MSILocation%\Administrator.msi" /Q
    if exist "%MSILocation%\CosmoDesigner.msi"                 DEL "%MSILocation%\CosmoDesigner.msi" /Q

:NODEL
    if not exist "%MSILocation%\CCSPClientInstallationService.msi" echo Downloading CCSPClientInstallationService.msi && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPClientInstallationService.msi' %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/ClientInstallationService/CCSPClientInstallationService.msi"
    if not exist "%MSILocation%\CCSPClientReportingService.msi"    echo Downloading CCSPClientReportingService.msi    && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPClientReportingService.msi'    %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPClientReportingService.msi"
    if not exist "%MSILocation%\CCSPClientCommunicator.msi"        echo Downloading CCSPClientCommunicator.msi        && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPClientCommunicator.msi'        %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPClientCommunicator.msi"
    if not exist "%MSILocation%\CCSPClientTrayApp.msi"             echo Downloading CCSPClientTrayApp.msi             && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPClientTrayApp.msi'             %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPClientTrayApp.msi"
    if not exist "%MSILocation%\CCSPClientUploadsService.msi"      echo Downloading CCSPClientUploadsService.msi      && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPClientUploadsService.msi'      %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPClientUploadsService.msi"
    if not exist "%MSILocation%\CCSPScreenRecordingService.msi"    echo Downloading CCSPScreenRecordingService.msi    && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPScreenRecordingService.msi'    %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPScreenRecordingService.msi"
    if not exist "%MSILocation%\CCSPSIPService.msi"                echo Downloading CCSPSIPService.msi                && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPSIPService.msi'                %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPSIPService.msi"
    if not exist "%MSILocation%\CCSPTouchPointConnector.msi"       echo Downloading CCSPTouchPointConnector.msi       && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CCSPTouchPointConnector.msi'       %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/CCSPTouchPointConnector.msi"
    if not exist "%MSILocation%\Encoder_en.exe"                    echo Downloading Encoder_en.exe                    && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\Encoder_en.exe'                    %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/Encoder_en.exe"
    if not exist "%MSILocation%\vcredist_x86.exe"                  echo Downloading vcredist_x86.exe                  && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\vcredist_x86.exe'                  %protocol%://%TouchPointHOSTNAME%/TouchPoint/ClientServices/vcredist_x86.exe"
    if not exist "%MSILocation%\Administrator.msi"                 echo Downloading Administrator.msi                 && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\Administrator.msi'                 %protocol%://%AdminHOSTNAME%/webadministrator/administrator.msi"
    if not exist "%MSILocation%\CosmoDesigner.msi"                 echo Downloading CosmoDesigner.msi                 && powershell.exe -Command "Invoke-WebRequest -OutFile '%MSILocation%\CosmoDesigner.msi'                 %protocol%://%AdminHOSTNAME%/cosmodesigner/cosmodesigner.msi"
    echo.
    echo.


:::: Add domain to compatibility view list. ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    echo Configuring IE Browser Settings .......................................................................................

:::: Add domain to compatibility view list. ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ::    note: requires running bat as administrator. The domain will not show in compatibility view list because it uses legacy group policy reg key method, however you can validate it is runnign in compatibliity view using F12.
    echo Adding %DOMAIN% to Compatibility Mode List
    REG ADD "HKCU\SOFTWARE\Policies\Microsoft\Internet Explorer\BrowserEmulation\PolicyList" /v "%DOMAIN%" /t REG_SZ /d "%DOMAIN%" /f
    echo.

:::: Add domain to Pop Up Blocker Allow list  ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    echo Adding %DOMAIN% to Pop Up Blocker Allow List
    REG ADD "HKCU\SOFTWARE\Microsoft\Internet Explorer\New Windows\Allow" /v "*.%DOMAIN%" /t REG_BINARY /d 000 /f
    echo.

:::: Add domain to Trusted Sites List and adjust Zone security settings ::::::::::::::::::::::::::::::::::
    ::     note 0=Enabled 1=Prompt 3=Disable
    echo Adding %DOMAIN% to Trusted Sites List
    REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\%DOMAIN%" /v * /t REG_DWORD /d 2 /f
    echo.
    echo Applying Custom Security Settings for Trusted Security Zone
    ::       Enable - Miscellaneous: Access data sources across domains (1406)
    ::       Enable - ActiveX controls and plug-ins: Download signed ActiveX controls (1001)
    ::       Prompt - ActiveX controls and plug-ins: Download unsigned ActiveX controls (1004)
    ::       Prompt - ActiveX controls and plug-ins: Initialize and script ActiveX controls not marked as safe for scripting (1201)

    REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2" /v 1406 /t REG_DWORD /d 0 /f
    REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2" /v 1001 /t REG_DWORD /d 0 /f
    REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2" /v 1004 /t REG_DWORD /d 1 /f
    REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2" /v 1201 /t REG_DWORD /d 1 /f
    echo.
    echo.


:::: Check for previously installed software :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::   This does a search of the install branch of the registry for the installed software names of the components.
::   for each object found, it finds the Display Name, Version, and the LocalPackage.  the local package is the cached msi windows uses to allow uninstall
::   this method is much faster than using wmic to call the uninstall

    echo Searching for previously installed software ...........................................................................
    for %%a in ("CosmoDesigner" "WebAdministrator" "CCSPClientInstallationService" "CCSPClientReportingService" "CCSPClientCommunicator" "CCSPClientTrayApp" "CCSPClientUploadsService" "CCSPScreenRecordingService" "CCSPSIPService" "CCSPTouchPointConnector") do (
       set SWNAME=%%~a                   
       for /f "tokens=*" %%b in ('reg QUERY HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\installer /f %%a /e /s ^| findstr /C:"HKEY_" ^| findstr /V /C:"Usage"') do (
          if not errorlevel 1 (
             set KEYNAME=%%b
             rem echo !KEYNAME!
             set KEYNAME=!KEYNAME:Features=InstallProperties!
             set KEYNAME=!KEYNAME:Usage=InstallProperties!
             rem echo !KEYNAME!

             rem FOR /F "tokens=3*" %%c IN ('REG Query "!KEYNAME!" /F DisplayVersion /V /E ^| FIND /I " DisplayVersion "')  DO (
             rem     set SWVER=%%c           
             rem     rem echo !SWVER!
             rem )
             
             FOR /F "tokens=3*" %%c IN ('REG Query "!KEYNAME!" /F DisplayName /V /E ^| FIND /I " DisplayName "')  DO (
                 set SWDISPLAY=%%c                      
                 rem echo !SWDISPLAY!
             )
             
             FOR /F "tokens=3*" %%c IN ('REG Query "!KEYNAME!" /F LocalPackage /V /E ^| FIND /I " LocalPackage "')  DO (
                 set MSIFILE=%%c
                 rem echo !MSIFILE!
             )
             rem echo Uninstalling !SWDISPLAY:~0,30! !SWVER:~0,12! !MSIFILE!
             echo Uninstalling !SWDISPLAY:~0,30!
             msiexec /x "!MSIFILE!" /passive !MSILOGGING!
             if !ERRORLEVEL! NEQ 0 echo Uninstall Failed Error !ERRORLEVEL!
          )
       )
    )
    echo.
    echo.


:::: 7.2 version prerequisites ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    REM  Installation of pre-requisites is only required if this has not already been done.
    echo Checking pre-requisites ...............................................................................................

    reg QUERY HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\installer /f "Microsoft Visual C++ 2008 Redistributable - x86 9.0.30729.17" /e /s > nul
    if %ERRORLEVEL% EQU 1 (
        echo Installing Microsoft Visual C++ 2008 Redistributable - x86
       "!MSILocation!vcredist_x86.exe" /q !VCRedistLog!
    )  else (
       echo Already Installed - Visual C++ 2008 Redistributable - x86
    )

    reg QUERY HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\installer /f "Microsoft Expression Encoder 4" /e /s > nul
    if %ERRORLEVEL% EQU 1 (
       echo Installing Microsoft Expression Encoder 4
       "!MSILocation!Encoder_en.exe" -q -r !EncoderLog!
    )  else (
       echo Already Installed - Expression Encoder 4
    )
    echo.
    echo.

:::: Install CCSP MSI files ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    echo Installing CCSP client software .......................................................................................
    echo Installing CCSPClientInstallationService
    msiexec /i "%MSILocation%\CCSPClientInstallationService.msi" /passive %MSILOGGING% USING_HTTPS=%usingSSL% INSTALLATIONCONNECTORURL="%protocol%://localhost:49074"
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing CCSPClientReportingService
    msiexec /i "%MSILocation%\CCSPClientReportingService.msi" /passive %MSILOGGING% usesWCF=True
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing CCSPClientCommunicator
    msiexec /i "%MSILocation%\CCSPClientCommunicator.msi" /passive %MSILOGGING% TOUCHPOINTCONNECTORURL="%protocol%://localhost:49071"
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing CCSPClientTrayApp
    msiexec /i "%MSILocation%\CCSPClientTrayApp.msi" /passive %MSILOGGING% TOUCHPOINTCONNECTORURL="%protocol%://localhost:49071"
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing CCSPClientUploadsService
    msiexec /i "%MSILocation%\CCSPClientUploadsService.msi" /passive %MSILOGGING%
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing CCSPScreenRecordingService
    msiexec /i "%MSILocation%\CCSPScreenRecordingService.msi" /passive %MSILOGGING% TOUCHPOINTCONNECTORURL="%protocol%://localhost:49071"
    if !ERRORLEVEL! NEQ 0 echo Install !!!!!! Error !ERRORLEVEL! !!!!!!

    echo Installing CCSPSIPService
    msiexec /i "%MSILocation%\CCSPSIPService.msi" /passive %MSILOGGING% SIPCONNECTORURL="%protocol%://localhost:49073" WCF_HTTP_PORT=%wcfHttpPort% WCF_HTTPS_PORT=%wcfHttpsPort%
    if !ERRORLEVEL! NEQ 0 echo !!!!!! Error !ERRORLEVEL! !!!!!!

    echo Installing CCSPTouchPointConnector
    msiexec /i "%MSILocation%\CCSPTouchPointConnector.msi" /passive %MSILOGGING% TOUCHPOINTCONNECTORURL="%protocol%://localhost:49071"
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing WebAdministrator
    msiexec /i "%MSILocation%\Administrator.msi" /passive %MSILOGGING% 
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo Installing CosmoDesigner
    msiexec /i "%MSILocation%\CosmoDesigner.msi" /passive %MSILOGGING%
    if !ERRORLEVEL! NEQ 0 echo    Installation reported an error !ERRORLEVEL! 

    echo.
    echo.

    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPClientInstallationService.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPClientReportingService.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPClientCommunicator.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPClientTrayApp.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPClientUploadsService.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPScreenRecordingService.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPSIPService.msi" /Q 
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPTouchPointConnector.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\CCSPWebAdministrator.msi" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\vcredist_x86.exe" /Q
    if "%UseWindowsTemp%" == "True" DEL "%MSILocation%\Encoder_en.exe" /Q



:::: Creating Application Shortcuts ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    echo Creating TouchPoint Application Shortcuts .............................................................................
    if exist "C:\Program Files\internet explorer\iexplore.exe"             echo Creating TouchPoint IExplorer Shortcut
    if exist "C:\Program Files\internet explorer\iexplore.exe"             powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%CD%\TouchPoint IExplorer.lnk'); $s.TargetPath='C:\Program Files\internet explorer\iexplore.exe';            $s.Arguments = ' %protocol%://%TouchPointHOSTNAME%/TouchPoint';                   $s.WorkingDirectory = 'C:\Program Files\internet explorer';              $s.Save()"

    if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" echo Creating TouchPoint Chrome Shortcut
    if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%CD%\TouchPoint Chrome.lnk');    $s.TargetPath='C:\Program Files (x86)\Google\Chrome\Application\chrome.exe';$s.Arguments = ' --new-window %protocol%://%TouchPointHOSTNAME%/TouchPoint';      $s.WorkingDirectory = 'C:\Program Files (x86)\Google\Chrome\Application';$s.Save()"

    if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" echo Creating TouchPoint Chrome App Shortcut
    if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%CD%\TouchPoint Chrome App.lnk');$s.TargetPath='C:\Program Files (x86)\Google\Chrome\Application\chrome.exe';$s.Arguments = ' --new-window --app=%protocol%://%TouchPointHOSTNAME%/TouchPoint';$s.WorkingDirectory = 'C:\Program Files (x86)\Google\Chrome\Application';$s.Save()"
    echo.
    echo.


echo Installation complete.
echo.
echo Press any key to close the window
pause > nul 2>&1
exit

