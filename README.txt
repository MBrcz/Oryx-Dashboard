MIT License
Copyright (c) Microsoft Corporation. All rights reserved.
Copyright (c) Python Software Foundation. All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE

-------------------------------------
Date: 01.01.2024
-------------------------------------

Attack on Europe 2022

In short:
The idea is to create to the dashboard that will utilize data collected by Oryx Blog that contains data of visually confirmed losses in the Second Russo-Ukrainian War in specified points in time (right now every quarter till start of conflict and when wayback machine started visiting Oryx blog).

Used technologies:
Python - Python Software Foundation; (for Web Scrapper - tool for gathering data),
Excel, Power Bi - Microsoft Corporation, Inc. (for the report).

The project consists of 4 directories:
a) ~\Data Source - the result of web scrapper and it's operations. It is the base for data modeling.
b) ~\Report - the directory that contains the report in .pbix and .pdf. In order to run the .pbix file, one need Microsoft Power Bi software - fortunatelly it is for free to download from: https://powerbi.microsoft.com/pl-pl/downloads/ (Polish version).
c)  ~\staticImg - Images as files used in project (dynamic images are saved as URL's and are in Data Source dir)
d) ~\Web Scrapper - the simple python script that is runned via file "Execute.bat". It is not compiled, so in order to run it one need a python 3.11 in order to set venv (virtual enviroment). Alternatively, you can use cmd to run it via main.py (do not forget resetting the venv and reinstalling the packages from requirements.txt!).

Data is scrapped from the Url(s):
https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html -> Russian Losses
https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html -> Ukrainian Losses

The dashboard is in Report/Oryx Report.pbix. It needs Microsoft Power Bi software in order to run, fortunatelly it is free to use and available to download from: 
https://powerbi.microsoft.com/en-us/downloads/

