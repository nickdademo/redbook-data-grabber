What?
=====
A graphical Python application which can automatically grab vehicle data (e.g. price, specifications, features etc.) from the Australian automotive website http://redbook.com.au. Vehicle images as well as page screenshots can also be grabbed. This data can then be exported to an Excel file.

Screenshot 1: https://raw.github.com/nickdademo/redbook-data-grabber/master/Screenshot1.png

Why?
====
This application was a personal project of mine and here were some of my motivations for creating it:  
1. To learn to create graphical applications using PyQt  
2. To learn more about automated web browsing and data grabbing, in particular using multiple threads to speed up the process  
3. To see if it was possible!

How?
====
This application uses PhantomJS (a headless browser) in combination with Selenium (a web automation framework) to automatically grab the data. PyQt provides a neat and easy-to-use user-interface for the application.

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
_The following procedure was tested with:_  
Ubuntu 12.04 LTS (32-bit)  
Python 2.7.3  
Selenium 2.39.0  
PhantomJS 1.9.2  

1. Install PIP:  
_$ sudo apt-get install python-pip_

2. Install Selenium Python bindings:  
_$ sudo pip install -U selenium_

3. Install BeautifulSoup:  
_$ sudo pip install beautifulsoup4_

4. Install html5lib (BeautifulSoup parser):  
_$ sudo pip install html5lib_

5. Install PyQt4:  
_$ sudo apt-get install python-qt4_

6. Install XlsxWriter:  
_$ sudo pip install XlsxWriter_

7. Download the latest version of PhantomJS from http://phantomjs.org/. Place the phantomjs binary executable in the same folder as the rdbg.py script.

8. Run application:  
_$ python rbdg.py_
