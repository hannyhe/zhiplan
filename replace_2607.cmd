@echo off
echo This script needs to run as Administrator to replace 2607.xlsx
echo.
echo It will:
echo 1. Stop Windows Search service temporarily
echo 2. Replace 2607.xlsx with 2607_new.xlsx
echo 3. Restart Windows Search
echo.
pause

net stop WSearch
if errorlevel 1 (
    echo Cannot stop WSearch. Please close any programs using 2607.xlsx and try again.
    pause
    exit /b 1
)

copy /Y "G:\me\????\2025\2607_new.xlsx" "G:\me\????\2025\2607.xlsx"
if errorlevel 1 (
    echo Copy failed.
    net start WSearch
    pause
    exit /b 1
)

del "C:\Users\coldt\AppData\Local\Temp\2607_original_bak.xlsx"
echo Replace successful!
net start WSearch
pause
