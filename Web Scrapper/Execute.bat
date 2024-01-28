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
