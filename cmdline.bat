@echo off
setlocal

:: Create output folder if it doesn't exist
if not exist "C:\Temp\CmdLines" mkdir "C:\Temp\CmdLines"

echo Collecting System Information...
systeminfo > "C:\Temp\CmdLines\sysinfo.txt"

echo Collecting DEP Settings...
wmic OS Get DataExecutionPrevention_SupportPolicy > "C:\Temp\CmdLines\dep.txt"

echo Listing Installed OnBase Applications...
wmic product list brief > "C:\Temp\CmdLines\installed_products.txt"

echo Checking Disk Capacity...
wmic logicaldisk > "C:\Temp\CmdLines\logicaldisk.txt"

echo CPU Usage (Manual placeholder, needs clarification)...
echo CPU Usage collection not defined > "C:\Temp\CmdLines\cpu_usage.txt"

echo Installed Updates...
wmic qfe list > "C:\Temp\CmdLines\installed_updates.txt"

echo Exporting Application Event Logs...
wevtutil qe Application /rd:true /f:Text "/q:*[System [(Level=2)]]" /c:100 > "C:\Temp\CmdLines\applog.txt"

echo Exporting Security Event Logs...
wevtutil qe Security /rd:true /f:Text "/q:*[System [(Level=2)]]" /c:100 > "C:\Temp\CmdLines\seclog.txt"

echo Exporting System Event Logs...
wevtutil qe System /rd:true /f:Text "/q:*[System [(Level=2)]]" /c:100 > "C:\Temp\CmdLines\syslog.txt"

echo Listing Windows Services...
wmic service list brief > "C:\Temp\CmdLines\services.txt"

echo Listing Web Sites and Application Pools...
%systemroot%\system32\inetsrv\AppCmd.exe list apppool > "C:\Temp\CmdLines\apppools.txt"

echo Listing Recycle Schedules...
%systemroot%\system32\inetsrv\AppCmd.exe list apppool "AppServer" /config > "C:\Temp\CmdLines\recycle_schedule.txt"

echo Listing Authentication Configuration...
%systemroot%\system32\inetsrv\AppCmd.exe list CONFIG -section:Authentication > "C:\Temp\CmdLines\authentication.txt"

echo Listing Session State Timeout...
%systemroot%\system32\inetsrv\AppCmd.exe list CONFIG "Default Web Site/AppServer" -section:system.web/sessionState > "C:\Temp\CmdLines\session_state.txt"

echo All tasks completed. Files saved in C:\Temp\CmdLines
pause
endlocal
