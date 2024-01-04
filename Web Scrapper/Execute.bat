@echo off
echo Restarting the virtual enviroment...
call python.exe -m venv venv

:: Set the paths to the virtual enviroments
set venv_deactivate=venv\Scripts\deactivate.bat
set venv_activate=venv\Scripts\activate.bat

:: Activate the virtual environment
echo Activating venv
call %venv_activate%

:: Redownloading packages
pip install -r requirements.txt

:: Calls script
echo Executing file: main.py
call python main.py

::Deactivating venv
echo Deactivation of venv
call %venv_deactivate%

:: Question Section.
setlocal enabledelayedexpansion

:ask_question
set /p "response=Do you want to open the report file? Do not forget to refresh (Y/N): "
if /i "!response!"=="Y" (
    echo The report file has been opened.
    cd /d ..\Report
    start "Report File" "Oryx Report.pbix"
) else if /i "!response!"=="N" (
    echo Fine, have a nice day!
) else (
    echo Please enter 'Y' or 'N'.
    goto ask_question
)
