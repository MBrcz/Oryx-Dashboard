-------------------------------------
Date: 01.01.2024
-------------------------------------
UPDATED: 16/02/2024


Attack on Europe 2022

In short:
The idea is to create to the dashboard that will utilize data collected by Oryx Blog that contains data of visually confirmed losses in the Second Russo-Ukrainian War in specified points in time (right now every quarter till start of conflict and when wayback machine started visiting Oryx blog).

The project consists of 4 directories:
a) ~\Data Source - the result of web scrapper and it's operations. It is the base for data modeling.
b) ~\Report - the directory that contains the report in .pbix and .pdf. In order to run the .pbix file, one need Microsoft Power Bi software - fortunatelly it is for free to download from: https://powerbi.microsoft.com/pl-pl/downloads/ (Polish version).
c)  ~\staticImg - Images as files used in project (dynamic images are saved as URL's and are in Data Source dir)
d) ~\Web Scrapper - the simple python script that is runned via file "Execute.bat". It is not compiled, so in order to run it one need a python 3.11 in order to set venv (virtual enviroment). Alternatively, you can use cmd to run it via main.py (do not forget resetting the venv and reinstalling the packages from requirements.txt!).

Data is scrapped from the Url(s):
https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html -> Russian Losses
https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html -> Ukrainian Losses

Used technologies:
Python - Python Software Foundation; (for Web Scrapper - tool for gathering data),
Ms Excel, Power Bi, Power Query, Dax, Power Pivot - Microsoft Corporation, Inc. (for the report).


------------------
Update (16/02/2024):
Added the version of the project made in Ms Excel. In order for project to work - user must enable macros (in other cases user will be soft blocked in first page only).
