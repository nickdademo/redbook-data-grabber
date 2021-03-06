What?
=====
A graphical Python application which can automatically grab vehicle data (e.g. price, specifications, features etc.) from the Australian automotive website http://redbook.com.au. Vehicle images as well as page screenshots can also be grabbed. This data can then be exported to an Excel file.

Screenshot 1: https://raw.github.com/nickdademo/redbook-data-grabber/master/Screenshot1.png

Why?
====
This application was a personal project of mine and here were some of my motivations for creating it:  
1. To learn to create graphical applications using PyQt.  
2. To learn more about automated web browsing and data grabbing, in particular, using multiple threads to speed up the process.  
3. To see if it was possible and to show others how!  

How?
====
This application makes use of the following:  
- PhantomJS, a headless browser: http://phantomjs.org/  
- Selenium, a web automation framework: http://docs.seleniumhq.org/  
- PyQt, a Python binding of the cross-platform GUI toolkit Qt: http://www.riverbankcomputing.com/software/pyqt/  

The application can be run on both Windows and Linux.

DISCLAIMER
==========
This application has been created and published for EDUCATIONAL purposes only!

Before using this program, you must first agree to RedBook.com.au's Terms & Conditions: http://www.redbook.com.au/help/terms-conditions

According to #4: _"Use of this website is for your personal and non-commercial use only. Except for the material held in your computer’s cache or a single permanent copy of the material for your personal use, you must not: ..."_

**Therefore, you must only use this application for PERSONAL USE! I am not responsible for any misuse of this application.**

Usage Instructions
==================
Linux
-----
_The following procedure has been tested with:_  
Ubuntu 14.04 LTS (64-bit)  
Python 3.4.0  
Selenium 2.44.0  
PhantomJS 1.9.7  
PyQt4 4.10.4  
BeautifulSoup4 4.3.2  
html5lib 0.999  
XlsxWriter 0.6.6  

1. Install PIP:  
_$ sudo apt-get install python3-pip_

2. Install Selenium Python bindings:  
_$ sudo pip3 install selenium_

3. Install BeautifulSoup:  
_$ sudo pip3 install beautifulsoup4_

4. Install html5lib (BeautifulSoup parser):  
_$ sudo pip3 install html5lib_

5. Install PyQt4:  
_$ sudo apt-get install python3-pyqt4_

6. Install XlsxWriter:  
_$ sudo pip3 install XlsxWriter_

7. Download the latest version of PhantomJS from http://phantomjs.org/. Place the phantomjs binary executable in the same folder as the rdbg.py script.

8. Run application:  
_$ python3 rbdg.py_

Windows
-------
_The following procedure has been tested with:_  
Windows 8.1 Professional (64-bit)  
Python 3.4.2  
Selenium 2.44.0  
PhantomJS 2.0.0  
PyQt4 4.11.3 for Py3.4 (x64) (Qt 5.3.2)  
BeautifulSoup4 4.3.2  
html5lib 0.999  
XlsxWriter 0.6.6  

1. Install Selenium Python bindings:  
_$ pip install selenium_

2. Install BeautifulSoup:  
_$ pip install beautifulsoup4_

3. Install html5lib (BeautifulSoup parser):  
_$ pip install html5lib_

4. Download and install PyQt4 from http://www.riverbankcomputing.com/software/pyqt/download.

5. Install XlsxWriter:  
_$ pip install XlsxWriter_

6. Download the latest version of PhantomJS from http://phantomjs.org/. Place the phantomjs binary executable in the same folder as the rdbg.py script.

7. Run application:  
_$ python rbdg.py_
