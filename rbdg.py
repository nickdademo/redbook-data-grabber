import sys
import traceback
import os
import time
import re
import string
from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import QThreadPool, QObject, QRunnable, QCoreApplication, pyqtSignal
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
import codecs
from collections import defaultdict
from bs4 import BeautifulSoup
import json
import urllib
import subprocess
from xlsxwriter.workbook import Workbook

################
# KNOWN ISSUES #
################
# 1: Used (All) Search Type is non-functional
# 2: The retry attempt immediately after a fail does not use the previously completed vehicle list

DEBUG                           = True
APP_VERSION_STRING              = "1.0.0"
DEFAULT_REDBOOK_URL             = "http://www.redbook.com.au/"
AUTHOR_STRING                   = "Created by Nick D'Ademo"
PHANTOMJS_PATH                  = "./phantomjs" # Change to phantomjs.exe if running in Windows. Also, "html5lib" is used as the BeautifulSoup parse library in Linux. For Windows, change line 528 to: BeautifulSoup(html_source)
WEBDRIVER_UNTIL_WAIT_S          = 30
WEBDRIVER_PAGELOAD_TIMEOUT_S    = 30
WEBDRIVER_IMPLICIT_WAIT_S       = 5
DEFAULT_SEARCH_TYPE_ID          = "rdbCurrent"
MAX_THREAD_COUNT                = 2
DATA_SAVE_REL_PATH              = "Data"
THREAD_STOP_WAIT_MS             = 1000
N_ROWS_GAP_BETWEEN_IMAGES       = 30
COLUMN_WIDTH_PADDING            = 2
EXCEL_TEMP_DIR                  = "Temp"

class Vehicle():
    def __init__(self, make, model, year, name, id, data_path):
        self.make = make
        self.model = model
        self.year = year
        self.name = name
        self.id = id
        self.data_path = data_path

class Error(Exception):
    """Base class for exceptions in this module."""
    pass

class StopException(Error):
    def __init__(self):
        pass

class WorkerSignals(QObject):
    result = pyqtSignal(list)
    retry = pyqtSignal(str)
    vehicle = pyqtSignal(object)
    log = pyqtSignal(str)

class WindowSignals(QObject):
    stopThreads = pyqtSignal()

def getValidSelectFromList(elements):
    for e in elements:
        if all(v.text != "" for v in Select(e).options):
            return Select(e)
    raise Exception("Could not get valid element from list.")

class ExportThread(QtCore.QThread):
    vehicleDone = pyqtSignal(str)

    def __init__(self, filename, timestamp, selectedVehicles, selectedParameters):
        super(ExportThread, self).__init__()
        self.filename = filename
        self.timestamp = timestamp
        self.selectedVehicles = selectedVehicles
        self.selectedParameters = selectedParameters
        
    def run(self):
        if not os.path.exists(EXCEL_TEMP_DIR):
            os.makedirs(EXCEL_TEMP_DIR)
        workbook = Workbook(str(self.filename), {'tmpdir': EXCEL_TEMP_DIR})
        # Add styles to workbook
        center_bold = workbook.add_format()
        center_bold.set_align('center')
        center_bold.set_align('vcenter')
        center_bold.set_bold()      
        center = workbook.add_format()
        center.set_align('center')
        center.set_align('vcenter')
        url_format = workbook.add_format({
            'color':     'blue',
            'underline': 1
        })
        url_format.set_align('center')
        url_format.set_align('vcenter')
        # Create dict
        field_count = defaultdict(int)
        # Lists to hold column widths
        param_col_0 = list()
        param_col_1 = list()
        param_col_2 = list()
        index_col_0 = list()
        index_col_1 = list()
        index_col_2 = list()
        index_col_3 = list()
        index_col_4 = list()
        index_col_5 = list()
        index_col_6 = list()
        col_0 = list()
        col_1 = list()
        col_2 = list()
        col_3 = list()
        # Create parameter name list
        p_list = [a.text(0) for a in self.selectedParameters]
        # Create parameter list sheet
        param_worksheet = workbook.add_worksheet("Parameter List")
        # Add index sheet
        j=0
        index_worksheet = workbook.add_worksheet("Index")
        index_worksheet.write(j, 0, "ID", center_bold)
        index_col_0.append(len("ID"))
        index_worksheet.write(j, 1, "Make", center_bold)
        index_col_1.append(len("Make"))
        index_worksheet.write(j, 2, "Model", center_bold)
        index_col_2.append(len("Model"))
        index_worksheet.write(j, 3, "Year", center_bold)
        index_col_3.append(len("Year"))
        index_worksheet.write(j, 4, "Full Name", center_bold)
        index_col_4.append(len("Full Name"))
        index_worksheet.write(j, 5, "Number of Images", center_bold)
        index_col_5.append(len("Number of Images"))
        index_worksheet.write(j, 6, "Price", center_bold)
        index_col_6.append(len("Price"))
        j += 1
        # Save each selected vehicle
        for v in self.selectedVehicles:
            year_item = v.parent()
            model_item = year_item.parent()
            make_item = model_item.parent()
            # File name
            file_path = v.data(0, QtCore.Qt.UserRole).toString()
            # Get ID
            m_obj = re.search(r"_(\d+)", file_path)
            if m_obj:
                id = m_obj.group(1) 
            else:
                raise Exception("Could not extract ID from JSON filename.")
            # Get name
            filename_noext = os.path.basename(str(file_path)).replace(".json", "")
            filename_noext_noid = os.path.basename(str(file_path)).replace("_" + str(id) + ".json", "")
            # Add to index
            index_worksheet.write_url(j, 0, "internal:" + "'" + str(id) + "'" + "!A1", url_format, str(id))
            index_col_0.append(len(str(id)))
            index_worksheet.write(j, 1, str(make_item.text(0)), center)
            index_col_1.append(len(str(make_item.text(0))))
            index_worksheet.write(j, 2, str(model_item.text(0)), center)
            index_col_2.append(len(str(model_item.text(0))))
            index_worksheet.write(j, 3, str(year_item.text(0)), center)
            index_col_3.append(len(str(year_item.text(0))))
            index_worksheet.write(j, 4, filename_noext_noid, center)
            index_col_4.append(len(filename_noext_noid))
            # Create sheet
            worksheet = workbook.add_worksheet(str(id))
            row = 0
            # Add link back to index
            worksheet.write_url(row, 0, "internal:" + "'" + "Index" + "'" + "!A1", url_format, "Back to Index")
            col_0.append(len("Back to Index"))
            row += 1
            # Get data
            with open(file_path) as data_file:    
                data = json.load(data_file)
            for key, value in data.iteritems(): 
                # Dictionary
                if isinstance(value, dict):
                    for k, v in value.iteritems():
                        if k == "Price":
                            index_worksheet.write(j, 6, v, center)
                            index_col_6.append(len(v))
                        if k in p_list:
                            if isinstance(v, dict):
                                # Add to count
                                field_count[k] += 1
                                for k2, v2 in v.iteritems():
                                    worksheet.write(row, 0, key, center_bold)
                                    col_0.append(len(key))
                                    worksheet.write(row, 1, k, center_bold)
                                    col_1.append(len(k))
                                    worksheet.write(row, 2, k2, center)
                                    col_2.append(len(k2))
                                    worksheet.write(row, 3, v2, center)
                                    col_3.append(len(v2))
                                    row += 1
                            else:
                                worksheet.write(row, 0, key, center_bold)
                                col_0.append(len(key))
                                worksheet.write(row, 1, k, center_bold)
                                col_1.append(len(k))
                                worksheet.write(row, 2, v, center)
                                col_2.append(len(v))
                                row += 1
                                # Add to count
                                field_count[k] += 1
                # List
                else:
                    for i in value:
                        if i in p_list:
                            worksheet.write(row, 0, key, center_bold)
                            col_0.append(len(key))
                            worksheet.write(row, 1, i, center)
                            col_1.append(len(i))
                            row += 1
                            # Add to count
                            field_count[i] += 1
            # Add images (if they exist)
            noMatches = True
            nImages = 0
            row += 1
            path_ = os.path.dirname(os.path.abspath(str(file_path)))
            for file in os.listdir(path_):
                # Vehicle image
                if re.match(re.sub(r'([()])', r'\\\1', filename_noext) + "_\d+\.(jpg|png|jpeg)$", file):
                    worksheet.insert_image(row, 0, os.path.abspath(path_ + "/" + file))
                    row += N_ROWS_GAP_BETWEEN_IMAGES
                    nImages += 1
                    noMatches = False
                # RedBook screenshot
                elif re.match(re.sub(r'([()])', r'\\\1', filename_noext) + "\.(jpg|png|jpeg)$", file):
                    worksheet.insert_image(0, 4, os.path.abspath(path_ + "/" + file))
                    noMatches = False
            # Print warning if no images were found for current vehicle
            if noMatches:
                print "WARNING: No images found for " + filename_noext
            index_worksheet.write(j, 5, str(nImages), center)
            index_col_5.append(len(str(nImages)))
            j += 1
            # Set column widths
            index_worksheet.set_column(0, 0, max(index_col_0) + COLUMN_WIDTH_PADDING)
            index_worksheet.set_column(1, 1, max(index_col_1) + COLUMN_WIDTH_PADDING)
            index_worksheet.set_column(2, 2, max(index_col_2) + COLUMN_WIDTH_PADDING)
            index_worksheet.set_column(3, 3, max(index_col_3) + COLUMN_WIDTH_PADDING)
            index_worksheet.set_column(4, 4, max(index_col_4) + COLUMN_WIDTH_PADDING)
            index_worksheet.set_column(5, 5, max(index_col_5) + COLUMN_WIDTH_PADDING)
            if len(col_0) > 0:
                worksheet.set_column(0, 0, max(col_0) + COLUMN_WIDTH_PADDING)
            if len(col_1) > 0:
                worksheet.set_column(1, 1, max(col_1) + COLUMN_WIDTH_PADDING)
            if len(col_2) > 0:
                worksheet.set_column(2, 2, max(col_2) + COLUMN_WIDTH_PADDING)
            if len(col_3) > 0:
                worksheet.set_column(3, 3, max(col_3) + COLUMN_WIDTH_PADDING)
            # Emit signal
            self.vehicleDone.emit(filename_noext_noid)
        # Add data to Parameter List sheet
        l=0
        for p in self.selectedParameters:
            param_worksheet.write(l, 0, str(p.parent().text(0)), center)
            param_col_0.append(len(str(p.parent().text(0))))
            param_worksheet.write(l, 1, str(p.text(0)), center)
            param_col_1.append(len(str(p.text(0))))
            param_worksheet.write(l, 2, field_count[str(p.text(0))], center)
            param_col_2.append(len(str(field_count[str(p.text(0))])))
            l+=1
        param_worksheet.set_column(0, 0, max(param_col_0) + COLUMN_WIDTH_PADDING)
        param_worksheet.set_column(1, 1, max(param_col_1) + COLUMN_WIDTH_PADDING)
        param_worksheet.set_column(2, 2, max(param_col_2) + COLUMN_WIDTH_PADDING)
        # Done
        workbook.close()
        # Remove temp dir
        try:
            os.rmdir(EXCEL_TEMP_DIR)
        except OSError:
            pass

class Worker(QRunnable):
    def __init__(self, url, searchTypeID, make, path, saveScreenshot, saveImages, completedVehicles):
        super(Worker, self).__init__()
        self.url = url
        self.searchTypeID = searchTypeID
        self.make = make
        self.path = path
        self.saveScreenshot = saveScreenshot
        self.saveImages = saveImages
        self.signals = WorkerSignals()
        self.setAutoDelete(True) # not needed if set to True
        self.doStop = False
        self.vehicleList = list()
        self.completedVehicles = completedVehicles
        # Get MAKE (without count)
        self.make_text = self.make
        m_obj = re.search(r"(.*) \(\d+\)", self.make)
        if m_obj:
            self.make_text = m_obj.group(1)   

    def run(self):
        doRetry = False
        try:
            # Check flag
            if self.doStop == False:
                # Start PhantomJS webdriver
                service_args = [
                    #'--proxy-type=none'
                    ]
                driver = webdriver.PhantomJS(PHANTOMJS_PATH, service_args=service_args)
                driver.set_page_load_timeout(WEBDRIVER_PAGELOAD_TIMEOUT_S)
                driver.implicitly_wait(WEBDRIVER_IMPLICIT_WAIT_S)
                driver.set_window_size(1024, 768) # optional
                # Iterate over MODELS
                modelIndex = 0
                while 1:
                    # Check flag
                    if self.doStop == True:
                        raise StopException()
                    # Open RedBook.com.au
                    driver.get(self.url)
                    # Wait for any Ajax requests to complete
                    WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda s: s.execute_script("return jQuery.active == 0"))
                    # Select search type
                    driver.find_element_by_id(self.searchTypeID).click()
                    # Wait for any Ajax requests to complete
                    WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda s: s.execute_script("return jQuery.active == 0"))
                    # Get valid MAKE select element
                    makes = driver.find_elements_by_id('cboMake')
                    m = getValidSelectFromList(makes)
                    # Select MAKE
                    makeClicked = False
                    for o in m.options:
                        if o.text == self.make:
                            o.click()
                            makeClicked = True
                            break
                    if not makeClicked:
                        raise Exception("Could not click select Make: " + self.make + " [Select element length=" + str(len(m.options)) + "]")
                    # Wait for any Ajax requests to complete
                    WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda s: s.execute_script("return jQuery.active == 0"))
                    # Get valid MODEL select element
                    models = driver.find_elements_by_id('cboModel')
                    model = getValidSelectFromList(models)
                    # Save number of models
                    nModels = len(model.options)
                    # Save current MODEL
                    model_option = model.options[modelIndex]
                    model_text = model_option.text
                    if model_text != "All Models" and model_option.get_attribute("source") == None:
                        # Select MODEL
                        model.select_by_visible_text(model_text)
                        # Get search button
                        search_buttons = driver.find_elements_by_id('btnSearch')
                        searchButtonClicked = False
                        for button in search_buttons:
                            if button.is_displayed():
                                searchButtonClicked = True
                                button.click()
                                break
                        if not searchButtonClicked:
                            raise Exception("Could not click Search button.")
                        # Iterate over RESULTS
                        resultIndex = 0
                        while 1:
                            # Check flag
                            if self.doStop == True:
                                raise StopException()
                            # Get list of ALL results (over multiple pages if applicable)
                            if resultIndex == 0:
                                # Wait for page to load
                                try:
                                    results = [x.get_attribute("href") for x in WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_elements_by_xpath("//div[@class='content']/a[@class='newcars']"))]
                                except Exception, e:
                                    # Do check
                                    no_results = WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_elements_by_xpath("//div[@class='']/div[@class='no-results']"))
                                    # Break out of loop
                                    break
                                 # Do we have multiple pages?
                                res_pages_links = driver.find_elements_by_xpath("//ul[@class='pagination']/li/a[text()!='Next']")
                                if res_pages_links != None:
                                    href_list = [a.get_attribute("href") for a in res_pages_links]
                                # Follow links
                                for href in href_list:
                                    driver.get(href)
                                    # Wait for page to load and append to results list
                                    results.append([x.get_attribute("href") for x in WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_elements_by_xpath("//div[@class='content']/a[@class='newcars']"))])    
                                # Save number of results
                                nResults = len(results)
                            # Get link
                            result = results[resultIndex]
                            # Save year from link href attribute
                            m_obj = re.search(r"/(\d{4})$", result)
                            if m_obj:
                                year = m_obj.group(1)
                            else:
                                raise Exception("Year attribute could not be extracted from element.")
                            # Follow link
                            driver.get(result)
                            
                            # Get list of ALL badges (over multiple pages if applicable)
                            # Wait for page to load
                            badges = dict((el.get_attribute("id"), el.get_attribute("href")) for el in WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_elements_by_xpath("//div[@class='content']/a[@class='item']")))
                            # Do we have multiple pages?
                            badges_pages_links = driver.find_elements_by_xpath("//ul[@class='pagination']/li/a[text()!='Next']")
                            if badges_pages_links != None:
                                href_list = [a.get_attribute("href") for a in badges_pages_links]
                            # Follow links
                            for href in href_list:
                                driver.get(href)
                                # Wait for page to load and append to results list
                                badges = dict(badges.items() + dict((el.get_attribute("id"), el.get_attribute("href")) for el in WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_elements_by_xpath("//div[@class='content']/a[@class='item']"))).items())
                            # Save number of badges        
                            nBadges = len(badges)
                            # Reset index
                            badgeIndex=0
                            # Iterate over BADGES
                            for id, href in badges.iteritems():
                                # Check flag
                                if self.doStop == True:
                                    raise StopException()
                                # Check if vehicle data has already been saved
                                if id not in self.completedVehicles:
                                    # Open page
                                    driver.get(href)
                                    # Wait for page to load
                                    badge = WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_element_by_xpath("//div[@class='details']/div/h1[@class='details-title']"))
                                    #############
                                    # SAVE DATA #
                                    #############
                                    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
                                    badge_text = badge.text
                                    # Create directory structure
                                    path_ = self.path + "/" + str(self.make_text) + "/" + str(model_text.strip()) + "/" + str(year)
                                    if not os.path.exists(path_):
                                        os.makedirs(path_)
                                    # Save data as JSON file
                                    filename = ''.join(c for c in badge_text if c in valid_chars)
                                    json_string = self.getDataAsJSONString(driver.page_source)
                                    file_path = path_ + "/" + filename + "_" + id + ".json"
                                    with codecs.open(file_path, mode="w", encoding='utf-8') as data_file:
                                        data_file.write(json_string)
                                    # Add to vehicle object list
                                    v = Vehicle(self.make_text, model_text.strip(), year, badge_text, id, file_path)
                                    self.vehicleList.append(v)
                                    # Save screenshot (if enabled)
                                    if self.saveScreenshot == True:
                                        self.takeScreenshot(driver, filename + "_" + id + '.png', path_)
                                    # Save images (if enabled)
                                    if self.saveImages == True:
                                        # Attempt to grab img elements
                                        try:
                                            images = None
                                            images = WebDriverWait(driver, WEBDRIVER_UNTIL_WAIT_S).until(lambda driver : driver.find_elements_by_xpath("//div/ul[@class='thumbs']/li/a/img"))
                                        # No images (timeout)
                                        except Exception, e:
                                            pass
                                        # Images found
                                        if images != None:
                                            i = 1
                                            for img in images:
                                                # Get URL
                                                m_obj = re.search(r"(.*)(\..*)\?", img.get_attribute("src"))
                                                if m_obj:
                                                    img_url_no_ext = m_obj.group(1)
                                                    img_ext = m_obj.group(2)
                                                else:
                                                    raise Exception("Image URL could not be extracted from element.")                                            
                                                urllib.urlretrieve(img_url_no_ext + img_ext, path_ + "/" + filename + "_" + id + "_" + str(i) + img_ext)
                                                i += 1
                                    # Emit signal
                                    self.signals.vehicle.emit(v)
                                    # Add to log
                                    self.signals.log.emit("Processed (" + id + "): " + badge_text)
                                    # Go back
                                    driver.back()
                                else:
                                    # Add to log
                                    self.signals.log.emit("Skipped (" + id + "): " + self.make_text)
                                # Increment index
                                badgeIndex+=1
                            # Go back
                            driver.back()
                            # Exit loop
                            resultIndex += 1
                            if resultIndex == nResults:
                                break
                    # Exit loop
                    modelIndex += 1
                    if modelIndex == nModels:
                        break
                # Done
                self.signals.result.emit(self.vehicleList)
            # Exit thread
            else:
                raise StopException()
        except StopException, e1:
            # Done
            self.signals.result.emit(self.vehicleList)
        except Exception, e2:
            if DEBUG:
                print traceback.format_exc()
            # Add to log
            self.signals.log.emit("Retrying: " + self.make_text)
            # Retry MAKE (set flag)
            doRetry = True
        finally:
            # Close webdriver (best effort)
            try:
                driver.close()
            except Exception, e3:
                pass
            finally:
                if doRetry:
                    self.signals.retry.emit(self.make)

    def takeScreenshot(self, driver, name, save_location):
        # Make sure the path exists
        path = os.path.abspath(save_location)
        if not os.path.exists(path):
            os.makedirs(path)
        full_path = "%s/%s" % (path, name)
        driver.get_screenshot_as_file(full_path)
        return full_path

    def getDataAsJSONString(self, html_source):
        # Parse
        #soup = BeautifulSoup(html_source) # For Windows
        soup = BeautifulSoup(html_source, "html5lib")
        # Initialize variable(s)
        json_string = ""
        fields = dict()

        # Valuation Prices: table[@id='ctl08_p_ctl04_ctl03_ctl01_ctl02_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl01_ctl02_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Check if value exists
                if value.findAll(text=True)[0] == None:
                    value_string = ""
                else:
                    value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Valuation Prices'] = temp

        # Overview: table[@id='ctl08_p_ctl04_ctl03_ctl02_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl02_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Exception #1: ANCAP Safety Rating, Overall Green Star Rating
                if label_string == "Overall Green Star Rating" or \
                label_string == "ANCAP Safety Rating":
                    for child in value.children:
                        class_string_1 = child['class'][1]
                        m_obj = re.search(r"[A-Z]*(\d{1,2})", class_string_1)
                        if m_obj:
                            # Single digit
                            if len(m_obj.group(1)) == 1:
                                value_string = m_obj.group(1)
                            # Double digit: add decimal point
                            else:
                                value_string = ".".join(m_obj.group(1))
                        else:
                            value_string = ""
                else:
                    # Check if value exists
                    if value.findAll(text=True)[0] == None:
                        value_string = ""
                    else:
                        value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Overview'] = temp

        # Engine: table[@id='ctl08_p_ctl04_ctl03_ctl05_ctl01_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl05_ctl01_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Check if value exists
                if value.findAll(text=True)[0] == None:
                    value_string = ""
                else:
                    value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Engine'] = temp

        # Dimensions: table[@id='ctl08_p_ctl04_ctl03_ctl05_ctl02_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl05_ctl02_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Check if value exists
                if value.findAll(text=True)[0] == None:
                    value_string = ""
                else:
                    value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Dimensions'] = temp

        # Warranty: table[@id='ctl08_p_ctl04_ctl03_ctl05_ctl03_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl05_ctl03_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Check if value exists
                if value.findAll(text=True)[0] == None:
                    value_string = ""
                else:
                    value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Warranty'] = temp

        # Green Info: div[@id='ctl08_p_ctl04_ctl03_ctl05_ctl04_pnlBody']
        temp = dict()
        for value in soup.select('div[id="ctl08_p_ctl04_ctl03_ctl05_ctl04_pnlBody"] tbody tr td[class*="definition"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Exception #1: Green Star Rating
                if "greenStarRating" in value['class']:
                    for child in value.children:
                        class_string_1 = child['class'][1]
                        m_obj = re.search(r"[A-Z]*(\d{1,2})", class_string_1)
                        if m_obj:
                            # Single digit
                            if len(m_obj.group(1)) == 1:
                                value_string = m_obj.group(1)
                            # Double digit: add decimal point
                            else:
                                value_string = ".".join(m_obj.group(1))
                        else:
                            value_string = ""
                else:
                    # Check if value exists
                    if value.findAll(text=True)[0] == None:
                        value_string = ""
                    else:
                        value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Green Info'] = temp

        # Steering: table[@id='ctl08_p_ctl04_ctl03_ctl05_ctl05_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl05_ctl05_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Check if value exists
                if value.findAll(text=True)[0] == None:
                    value_string = ""
                else:
                    value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Steering'] = temp

        # Wheels: table[@id='ctl08_p_ctl04_ctl03_ctl05_ctl06_dgProps']
        temp = dict()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl05_ctl06_dgProps"] td[class="value"]'):
            # Initialize variable(s)
            value_string = ""
            # Get label
            label_string = value.find_previous_sibling().string.strip()
            # Ignore empty label
            if label_string != "":
                # Check if value exists
                if value.findAll(text=True)[0] == None:
                    value_string = ""
                else:
                    value_string = value.findAll(text=True)[0].strip()
                # Add to temporary dict
                temp[label_string] = value_string
        # Add to dict
        if len(temp) > 0:
            fields['Wheels'] = temp

        # Standard Equipment: table[@id='ctl08_p_ctl04_ctl03_ctl05_ctl07_dgPropsNoLabel']
        temp = list()
        for value in soup.select('table[id="ctl08_p_ctl04_ctl03_ctl05_ctl07_dgPropsNoLabel"] tbody tr td[class*="item"]'):
            # Add to temporary dict
            temp.append(value.string.strip())
        # Add to dict
        if len(temp) > 0:
            fields['Standard Equipment'] = temp

        # Optional Features
        temp = dict()
        for value in soup.select('div[class*="optional-features"] div[class="csn-properties"]'):
            for div in value.children:
                if 'header' in div['class']:
                    # Make sure Heading is not empty
                    if div.find_all(text=True)[0] != None:
                        # Save heading string
                        heading_string = value.find_all(text=True)[0].strip()
                        # Get Body content
                        body_items = value.find("div","body").find_all('label')
                        # Create temp dict
                        temp2 = dict()
                        # Loop through items
                        for item in body_items:
                            # Label element has NO children
                            if item.string != None and item.find_all(text=True)[0] != None:
                                definition = item.find_parents("td","term")[0].find_next_sibling()
                                # Check if definition value exists
                                if definition.findAll(text=True)[0] == None:
                                    def_string = ""
                                else:
                                    def_string = definition.findAll(text=True)[0].strip()
                                temp2[item.find_all(text=True)[0].strip()] = def_string
                            # Label element has children
                            else:
                                for c in item.children:
                                    if c.find_all(text=True)[0] != None and c.find_all(text=True)[0].strip() != "":
                                        # Add price data if present (2nd text element)
                                        if len(c.find_all(text=True))>1 and c.find_all(text=True)[1] != None and c.find_all(text=True)[1].strip() != "": 
                                            temp2[c.find_all(text=True)[0].strip()] = c.find_all(text=True)[1].strip()
                                        # No price data
                                        else:
                                            temp2[c.find_all(text=True)[0].strip()] = ""
                        # Add to temporary dict
                        temp[heading_string] = temp2
        # Add to dict
        if len(temp) > 0:
            fields['Optional Features'] = temp

        # Return JSON string
        return json.dumps(fields, sort_keys=True, indent=4)

    def stop(self):
        self.doStop = True

class Window(QtGui.QWidget):
    def __init__(self):
        QtGui.QWidget.__init__(self)
        # Initialize UI
        self.initUI()
        # Create thread pool
        self.pool = QThreadPool()
        self.pool.setMaxThreadCount(MAX_THREAD_COUNT)
        # Create signals
        self.signals = WindowSignals()
        # Create dict
        self.completedVehicles = defaultdict(list)
        # Create dict of sets (holds all available fields)
        self.availableFields = dict()
        # Set thread to None
        self.exportThread = None
        # Check for data (load if present)
        if os.path.exists(DATA_SAVE_REL_PATH):
            folder_list = [os.path.abspath(DATA_SAVE_REL_PATH + "/" + name) for name in os.listdir(DATA_SAVE_REL_PATH)]
            if len(folder_list) > 0:
                # Set UI
                self.textedit_dataPath.setText(max(folder_list))
                # Do load
                self.loadData()
                # Show data
                self.showData()

    def closeEvent(self, event):
        if self.exportThread != None and self.exportThread.isRunning():
            QtGui.QMessageBox.warning(self, "Error", "Please wait for the Excel export process to finish before attempting to close the program.")
            event.ignore()
        elif not self.pushbutton_getDataStart.isEnabled():
            QtGui.QMessageBox.warning(self, "Error", "Please stop the get data process before attempting to close the program.")
            event.ignore()
        else:
            quit_msg = "Are you sure you want to exit the program?"
            reply = QtGui.QMessageBox.question(self, 'Message', quit_msg, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
            if reply == QtGui.QMessageBox.Yes:
                event.accept()
            else:
                event.ignore()

    def initUI(self):
        # Window
        self.resize(640, 800)
        self.center()
        self.setWindowTitle('RedBook Data Grabber' + ' ' + APP_VERSION_STRING + " - " + AUTHOR_STRING)
        # Label (URL)
        self.label_url = QtGui.QLabel("URL:")
        self.label_url.setMaximumHeight(25)
        # Text Edit (URL)
        self.url = QtGui.QTextEdit()
        self.url.setMaximumHeight(25)
        self.url.setText(DEFAULT_REDBOOK_URL)
        # Label (Search Type)
        self.label_searchtype = QtGui.QLabel("Search Type:")
        self.label_searchtype.setMaximumHeight(25)
        # Radio Button (Search Type)
        self.searchtype_group = QtGui.QButtonGroup()
        self.r0=QtGui.QRadioButton("Used (All)")
        self.r0.setMaximumHeight(25)
        if DEFAULT_SEARCH_TYPE_ID == 'rdbUsed':
            self.r0.setChecked(True)
        self.searchtype_group.addButton(self.r0)
        self.r1=QtGui.QRadioButton("New (Current)")
        self.r1.setMaximumHeight(25)
        if DEFAULT_SEARCH_TYPE_ID == 'rdbCurrent':
            self.r1.setChecked(True)
        self.searchtype_group.addButton(self.r1)
        # Layout (URL)
        self.url_hbox = QtGui.QHBoxLayout()
        self.url_hbox.addWidget(self.label_url)
        self.url_hbox.addWidget(self.url)
        self.url_hbox.addWidget(self.label_searchtype)
        self.url_hbox.addWidget(self.r0)
        self.url_hbox.addWidget(self.r1)
        # Label (Get Data)
        self.label_getData = QtGui.QLabel("Get Data:")
        # Push Button (Start)
        self.pushbutton_getDataStart = QtGui.QPushButton("Start")
        self.pushbutton_getDataStart.released.connect (self.getData)
        # Push Button (Stop)
        self.pushbutton_getDataStop = QtGui.QPushButton("Stop")
        self.pushbutton_getDataStop.setEnabled(False)
        self.pushbutton_getDataStop.released.connect (self.stop)
        # Checkbox (Save Screenshots)
        self.checkbox_saveScreenshots = QtGui.QCheckBox("Save Screenshots")
        # Checkbox (Save Images)
        self.checkbox_saveImages = QtGui.QCheckBox("Save Images")
        # Layout
        self.getData_layout = QtGui.QHBoxLayout()
        self.getData_layout.addWidget(self.label_getData)
        self.getData_layout.addWidget(self.pushbutton_getDataStart, 1)
        self.getData_layout.addWidget(self.pushbutton_getDataStop)
        self.getData_layout.addWidget(self.checkbox_saveScreenshots)
        self.getData_layout.addWidget(self.checkbox_saveImages)
        # Label (Data Path)
        self.label_dataPath = QtGui.QLabel("Data Path:")
        # Text Edit (Data Path)
        self.textedit_dataPath = QtGui.QTextEdit()
        self.textedit_dataPath.setMaximumHeight(25)
        # Push Button (Load Data)
        self.pushbutton_loadData = QtGui.QPushButton("Load Data...")
        self.pushbutton_loadData.released.connect (self.loadDataFromUIPath)
        # Layout
        self.dataPath_layout = QtGui.QHBoxLayout()
        self.dataPath_layout.addWidget(self.label_dataPath)
        self.dataPath_layout.addWidget(self.textedit_dataPath)
        self.dataPath_layout.addWidget(self.pushbutton_loadData)
        # Tree Widget (Vehicles)
        self.treeWidget_vehicles = QtGui.QTreeWidget()
        self.treeWidget_vehicles.setColumnCount(2)
        self.treeWidget_vehicles.setHeaderLabels(["Vehicle","ID"])
        self.treeWidget_vehicles.setHeaderHidden(False)
        self.treeWidget_vehicles.itemChanged.connect (self.handleChanged)
        self.treeWidget_vehicles.itemSelectionChanged.connect (self.showData)
        self.treeWidget_vehicles.itemExpanded.connect (self.autoResizeVehicles)
        self.treeWidget_vehicles.itemCollapsed.connect (self.autoResizeVehicles)
        self.treeWidget_vehicles.itemDoubleClicked.connect (self.showInFolder)
        # Buttons
        self.pushbutton_vehiclesExpandAll = QtGui.QPushButton("Expand All")
        self.pushbutton_vehiclesExpandAll.released.connect (self.expandAllVehicles)
        self.pushbutton_vehiclesCollapseAll = QtGui.QPushButton("Collapse All")
        self.pushbutton_vehiclesCollapseAll.released.connect (self.collapseAllVehicles)
        self.pushbutton_vehiclesSelectAll = QtGui.QPushButton("Select All")
        self.pushbutton_vehiclesSelectAll.released.connect (self.selectAllVehicles)
        self.pushbutton_vehiclesDeselectAll = QtGui.QPushButton("Deselect All")
        self.pushbutton_vehiclesDeselectAll.released.connect (self.deselectAllVehicles)
        # Layout
        self.vehiclesButtons_layout = QtGui.QHBoxLayout()
        self.vehiclesButtons_layout.addWidget(self.pushbutton_vehiclesExpandAll)
        self.vehiclesButtons_layout.addWidget(self.pushbutton_vehiclesCollapseAll)
        self.vehiclesButtons_layout.addWidget(self.pushbutton_vehiclesSelectAll)
        self.vehiclesButtons_layout.addWidget(self.pushbutton_vehiclesDeselectAll)
        # Progress Bar (Get Data)
        self.progressBar = QtGui.QProgressBar()
        self.progressBar.setFormat("%p% (%v/%m)")
        self.progressBar.setAlignment(QtCore.Qt.AlignCenter)
        # Label
        self.label_data = QtGui.QLabel("")
        # Tree Widget (Data)
        self.treeWidget_data = QtGui.QTreeWidget()
        self.treeWidget_data.setColumnCount(2)
        self.treeWidget_data.setHeaderLabels(["Field","Value"])
        self.treeWidget_data.setHeaderHidden(False)
        self.treeWidget_data.itemChanged.connect (self.handleChanged)
        self.treeWidget_data.itemExpanded.connect (self.autoResizeData)
        self.treeWidget_data.itemCollapsed.connect (self.autoResizeData)
        # Buttons
        self.pushbutton_dataExpandAll = QtGui.QPushButton("Expand All")
        self.pushbutton_dataExpandAll.released.connect (self.expandAllData)
        self.pushbutton_dataCollapseAll = QtGui.QPushButton("Collapse All")
        self.pushbutton_dataCollapseAll.released.connect (self.collapseAllData)
        self.pushbutton_dataSelectAll = QtGui.QPushButton("Select All")
        self.pushbutton_dataSelectAll.released.connect (self.selectAllData)
        self.pushbutton_dataDeselectAll = QtGui.QPushButton("Deselect All")
        self.pushbutton_dataDeselectAll.released.connect (self.deselectAllData)
        # Layout
        self.dataButtons_layout = QtGui.QHBoxLayout()
        self.dataButtons_layout.addWidget(self.pushbutton_dataExpandAll)
        self.dataButtons_layout.addWidget(self.pushbutton_dataCollapseAll)
        self.dataButtons_layout.addWidget(self.pushbutton_dataSelectAll)
        self.dataButtons_layout.addWidget(self.pushbutton_dataDeselectAll)
        # Text Edit (Log)
        self.log = QtGui.QTextEdit()
        self.log.setReadOnly(True)
        # Button
        self.pushbutton_exportSelectedDataToExcel = QtGui.QPushButton("Exported Selected Data to Excel Workbook")
        self.pushbutton_exportSelectedDataToExcel.released.connect (self.exportToExcel)
        # Progress Bar (Export)
        self.progressBar_export = QtGui.QProgressBar()
        self.progressBar_export.setFormat("%p% (%v/%m)")
        self.progressBar_export.setAlignment(QtCore.Qt.AlignCenter)
        # OVERALL LAYOUT
        layout = QtGui.QVBoxLayout()
        layout.addLayout(self.url_hbox)
        layout.addLayout(self.getData_layout)
        layout.addLayout(self.dataPath_layout)
        layout.addWidget(self.treeWidget_vehicles)
        layout.addLayout(self.vehiclesButtons_layout)
        layout.addWidget(self.progressBar)
        layout.addWidget(self.label_data)
        layout.addWidget(self.treeWidget_data)
        layout.addLayout(self.dataButtons_layout)
        layout.addWidget(self.log)
        layout.addWidget(self.pushbutton_exportSelectedDataToExcel)
        layout.addWidget(self.progressBar_export)
        self.setLayout(layout)

    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def addChild(self, parent, column, title, data, addCheck):
        if data != None:
            data_ = QtCore.QVariant(str(data))
        else:
            data_ = data
        item = QtGui.QTreeWidgetItem(parent, title)
        item.setData(column, QtCore.Qt.UserRole, data_)
        if addCheck:
            item.setCheckState (column, QtCore.Qt.Unchecked)
        item.setExpanded (False)
        return item

    def handleChanged(self, item, column):
        if item.checkState(column) == QtCore.Qt.Checked:
            for x in range (0, item.childCount()):
                item.child(x).setCheckState (column, QtCore.Qt.Checked)
        elif item.checkState(column) == QtCore.Qt.Unchecked:
            for x in range (0, item.childCount()):
                item.child(x).setCheckState (column, QtCore.Qt.Unchecked)

    def getData(self):
        # Save timestamp
        self.timestamp = int(time.time())
        # Save path
        self.path = DATA_SAVE_REL_PATH + "/" + str(self.timestamp)
        # Set UI
        self.pushbutton_getDataStart.setEnabled(False)
        self.url.setEnabled(False)
        self.r0.setEnabled(False)
        self.r1.setEnabled(False)
        self.checkbox_saveScreenshots.setEnabled(False)
        self.checkbox_saveImages.setEnabled(False)
        self.textedit_dataPath.setEnabled(False)
        self.pushbutton_loadData.setEnabled(False)
        self.pushbutton_exportSelectedDataToExcel.setEnabled(False)
        self.treeWidget_vehicles.clear()
        self.log.clear()
        self.textedit_dataPath.setText(os.path.abspath(self.path))
        self.progressBar.reset()
        self.progressBar_export.reset()
        # Clear dict(s)
        self.completedVehicles.clear()
        self.availableFields.clear()
        # Show data
        self.showData()
        # Execute UI changes
        QCoreApplication.processEvents()
        # Start PhantomJS webdriver
        try:
            service_args = [
                #'--proxy-type=none'
                ]
            driver = webdriver.PhantomJS(PHANTOMJS_PATH, service_args=service_args)
            driver.set_page_load_timeout(WEBDRIVER_PAGELOAD_TIMEOUT_S)
            driver.implicitly_wait(WEBDRIVER_IMPLICIT_WAIT_S)
            driver.set_window_size(1024, 768) # optional
            # Get URL
            url = str(self.url.toPlainText())
            # Open RedBook.com.au
            driver.get(url)
            # Get search type
            if self.searchtype_group.checkedButton().text() == 'Used (All)':
                searchTypeID = "rdbUsed"
            elif self.searchtype_group.checkedButton().text() == 'New (Current)':
                searchTypeID = "rdbCurrent"
            # Select search type
            driver.find_element_by_id(searchTypeID).click()
            # Save MAKE select box
            makes = driver.find_elements_by_id("cboMake")
            # Get valid element
            make = getValidSelectFromList(makes)
            # Start scraping
            self.addToLog("Starting web scraping with maximum thread count: " + str(MAX_THREAD_COUNT))
            self.nMakeProcessed = 0
            self.nMakeTotal = 0
            for make in make.options:
                if make.text != "All Makes":
                    # Get MAKE (without count)
                    make_text = make.text
                    m_obj = re.search(r"(.*) \(\d+\)", make.text)
                    if m_obj:
                        make_text = m_obj.group(1)
                    # Create and start thread
                    worker = Worker(url, searchTypeID, make.text, self.path, self.checkbox_saveScreenshots.isChecked(), self.checkbox_saveImages.isChecked(), self.completedVehicles[make_text])
                    worker.signals.result.connect(self.newMake, QtCore.Qt.QueuedConnection)
                    worker.signals.retry.connect(self.retryMake, QtCore.Qt.QueuedConnection)
                    worker.signals.vehicle.connect(self.newVehicle, QtCore.Qt.QueuedConnection)
                    worker.signals.log.connect(self.addToLog, QtCore.Qt.QueuedConnection)
                    self.signals.stopThreads.connect(worker.stop)
                    self.pool.start(worker)
                    self.nMakeTotal += 1
            # Set progress bar range
            self.progressBar.setRange(0, self.nMakeTotal)
            self.progressBar.setValue(0)
            # Set stop button
            self.pushbutton_getDataStop.setEnabled(True)
        except Exception, e1:
            # Add to log
            self.addToLog("Could not start web scraping: " + str(e1))
            # Set UI
            self.pushbutton_getDataStart.setEnabled(True)
            self.pushbutton_getDataStop.setEnabled(False)
            self.url.setEnabled(True)
            self.r0.setEnabled(True)
            self.r1.setEnabled(True)
            self.checkbox_saveScreenshots.setEnabled(True)
            self.checkbox_saveImages.setEnabled(True)
            self.textedit_dataPath.setEnabled(True)
            self.pushbutton_loadData.setEnabled(True)
            self.pushbutton_exportSelectedDataToExcel.setEnabled(True)
            self.progressBar.reset()
        finally:
            # Close webdriver (best effort)
            try:
                driver.close()
            except Exception, e2:
                pass

    def newMake(self, vehicleList):
        # Add data to tree widget
        for v in vehicleList:
            # Reset flags
            makeExists = False
            modelExists = False
            yearExists = False
            vehicleExists = False
            # Does MAKE exist?
            for item_make in self.treeWidget_vehicles.findItems(v.make, QtCore.Qt.MatchExactly):
                makeExists = True
                make_child = item_make
                break
            # Add MAKE
            if not makeExists:
                make_child = self.addChild(self.treeWidget_vehicles.invisibleRootItem(), 0, [v.make], self.path + "\\" + v.make, True)
            # Does MODEL exist?
            for x in range (0, make_child.childCount()):
                if make_child.child(x).text(0) == v.model:
                    modelExists = True
                    model_child = make_child.child(x)
                    break
            # Add MODEL
            if not modelExists:
                model_child = self.addChild(make_child, 0, [v.model], self.path + "\\" + v.make + "\\" + v.model, True)
            # Does YEAR exist?
            for x in range (0, model_child.childCount()):
                if model_child.child(x).text(0) == v.year:
                    yearExists = True
                    year_child = model_child.child(x)
                    break
            # Add YEAR
            if not yearExists:
                year_child = self.addChild(model_child, 0, [v.year], self.path + "\\" + v.make + "\\" + v.model + "\\" + v.year, True)
            # Does vehicle exist?
            for x in range (0, year_child.childCount()):
                if year_child.child(x).text(0) == v.name and year_child.child(x).text(1) == v.id:
                    vehicleExists = True
                    break
            # Add vehicle
            if not vehicleExists:
                self.addChild(year_child, 0, [v.name, v.id], v.data_path, True)
                # Save available fields
                self.saveFields(v.data_path)
        
        # Vehicles in list
        if len(vehicleList)>0:
            # Sort items
            self.treeWidget_vehicles.sortItems(0, QtCore.Qt.AscendingOrder)
            # Show data
            self.showData()
        # Auto resize columns
        self.treeWidget_vehicles.resizeColumnToContents(0);
        self.treeWidget_vehicles.resizeColumnToContents(1);
        # Update progress bar
        self.nMakeProcessed += 1
        self.progressBar.setValue(self.nMakeProcessed)
        # On finish:
        if self.nMakeProcessed == self.nMakeTotal:
            # Set UI
            self.pushbutton_getDataStart.setEnabled(True)
            self.pushbutton_getDataStop.setEnabled(False)
            self.url.setEnabled(True)
            self.r0.setEnabled(True)
            self.r1.setEnabled(True)
            self.checkbox_saveScreenshots.setEnabled(True)
            self.checkbox_saveImages.setEnabled(True)
            self.textedit_dataPath.setEnabled(True)
            self.pushbutton_loadData.setEnabled(True)
            self.pushbutton_exportSelectedDataToExcel.setEnabled(True)

    def newVehicle(self, v):
        # Add to dict
        self.completedVehicles[v.make].append(v.id)

    def retryMake(self, make):
        # Get URL
        url = str(self.url.toPlainText())
        # Get search type
        if self.searchtype_group.checkedButton().text()=='Used (All)':
            searchTypeID = "rdbUsed"
        elif self.searchtype_group.checkedButton().text()=='New (Current)':
            searchTypeID = "rdbCurrent"
        # Get MAKE (without count)
        make_text = make
        m_obj = re.search(r"(.*) \(\d+\)", make)
        if m_obj:
            make_text = m_obj.group(1)
        # Create and start thread
        worker = Worker(url, searchTypeID, make, self.path, self.checkbox_saveScreenshots.isChecked(), self.checkbox_saveImages.isChecked(), self.completedVehicles[make_text])
        worker.signals.result.connect(self.newMake, QtCore.Qt.QueuedConnection)
        worker.signals.retry.connect(self.retryMake, QtCore.Qt.QueuedConnection)
        worker.signals.vehicle.connect(self.newVehicle, QtCore.Qt.QueuedConnection)
        worker.signals.log.connect(self.addToLog, QtCore.Qt.QueuedConnection)
        self.signals.stopThreads.connect(worker.stop)
        self.pool.start(worker, QtCore.QThread.NormalPriority)

    def addToLog(self, text):
        self.log.append(text)

    def stop(self):
        # Set UI
        self.pushbutton_getDataStop.setEnabled(False)
        # Execute UI changes
        QCoreApplication.processEvents()
        # Stop all threads
        self.signals.stopThreads.emit()
        # Wait
        while not self.pool.waitForDone(THREAD_STOP_WAIT_MS):
            self.signals.stopThreads.emit()

    def showData(self):
        sel = self.treeWidget_vehicles.selectedItems()
        vehicleSelected = False
        for s in sel:
            # Item is vehicle
            if s.data(0, QtCore.Qt.UserRole) != None and s.isSelected() and os.path.isfile(s.data(0, QtCore.Qt.UserRole).toString()):
                self.treeWidget_data.clear()
                file_path = s.data(0, QtCore.Qt.UserRole).toString()
                vehicleSelected = True
                # Show data
                with open(file_path) as data_file:    
                    data = json.load(data_file)
                for key, value in data.iteritems():
                    item = self.addChild(self.treeWidget_data.invisibleRootItem(), 0, [key], None, False)
                    # Dictionary
                    if isinstance(value, dict):
                        for k, v in value.iteritems():
                            if isinstance(v, dict):
                                item_ = self.addChild(item, 0, [k], None, False)
                                for k2, v2 in v.iteritems():
                                    self.addChild(item_, 0, [k2,v2], None, False)
                            else:
                                self.addChild(item, 0, [k,v], None, False)
                    # List
                    else:
                        for i in value:
                            self.addChild(item, 0, [i], None, False)
                # Auto resize columns
                self.treeWidget_data.resizeColumnToContents(0)
                self.treeWidget_data.resizeColumnToContents(1)
                break

        # Show all available data fields if vehicle is not selected
        if not vehicleSelected:
            self.treeWidget_data.clear()
            # Show all available fields
            self.showAllAvailableFields()
            size = sum(len(v) for v in self.availableFields.itervalues())
            self.label_data.setText("Showing all available fields" + " " + "(" + str(size) + ")")
            # Enable select/deselect buttons
            self.pushbutton_dataSelectAll.setEnabled(True)
            self.pushbutton_dataDeselectAll.setEnabled(True)
        # Show path to JSON file
        else:
            self.label_data.setText(file_path)
            # Disable select/deselect buttons
            self.pushbutton_dataSelectAll.setEnabled(False)
            self.pushbutton_dataDeselectAll.setEnabled(False)
        # Sort
        self.treeWidget_data.sortItems(0, QtCore.Qt.AscendingOrder)

    def autoResizeData(self, item):
        # Auto resize columns
        self.treeWidget_data.resizeColumnToContents(0)
        self.treeWidget_data.resizeColumnToContents(1)

    def autoResizeVehicles(self, item):
        # Auto resize columns
        self.treeWidget_vehicles.resizeColumnToContents(0);
        self.treeWidget_vehicles.resizeColumnToContents(1);

    def selectAllData(self):
        root = self.treeWidget_data.invisibleRootItem()
        for x in range (0, root.childCount()):
            root.child(x).setCheckState (0, QtCore.Qt.Checked)

    def deselectAllData(self):
        root = self.treeWidget_data.invisibleRootItem()
        for x in range (0, root.childCount()):
            root.child(x).setCheckState (0, QtCore.Qt.Unchecked)

    def selectAllVehicles(self):
        root = self.treeWidget_vehicles.invisibleRootItem()
        for x in range (0, root.childCount()):
            root.child(x).setCheckState (0, QtCore.Qt.Checked)

    def deselectAllVehicles(self):
        root = self.treeWidget_vehicles.invisibleRootItem()
        for x in range (0, root.childCount()):
            root.child(x).setCheckState (0, QtCore.Qt.Unchecked)

    def expandAllData(self):
        self.treeWidget_data.expandAll()
         # Auto resize columns
        self.treeWidget_data.resizeColumnToContents(0);
        self.treeWidget_data.resizeColumnToContents(1);       

    def collapseAllData(self):
        self.treeWidget_data.collapseAll()
         # Auto resize columns
        self.treeWidget_data.resizeColumnToContents(0);
        self.treeWidget_data.resizeColumnToContents(1);     

    def expandAllVehicles(self):
        self.treeWidget_vehicles.expandAll()
         # Auto resize columns
        self.treeWidget_vehicles.resizeColumnToContents(0);
        self.treeWidget_vehicles.resizeColumnToContents(1);     

    def collapseAllVehicles(self):
        self.treeWidget_vehicles.collapseAll()
         # Auto resize columns
        self.treeWidget_vehicles.resizeColumnToContents(0);
        self.treeWidget_vehicles.resizeColumnToContents(1);

    def loadDataFromUIPath(self):
        try:
            # Prompt user for path
            dialog = QtGui.QFileDialog(self)
            dialog.setDirectory(DATA_SAVE_REL_PATH)
            dialog.setFileMode(QtGui.QFileDialog.Directory)
            # Prompt user
            if dialog.exec_():
                # Get path
                for dir in dialog.selectedFiles():
                    dirname = os.path.basename(str(dir))
                    # Check folder name
                    m_obj = re.search(r"(\d{10})", dirname)
                    if m_obj:
                        self.timestamp = int(m_obj.group(1))
                    else:
                        raise Exception("Error extracting timestamp from top-level data directory name.")
                    # Set text edit
                    self.textedit_dataPath.setText(os.path.abspath(dir))
                    # Load data
                    self.loadData()
                    break
        except Exception, e:
            # Add to log
            self.addToLog("Could not load data: " + str(e))

    def loadData(self):
        try:
            # Save path
            self.path = str(self.textedit_dataPath.toPlainText())
            # Save timestamp
            self.timestamp = int(os.path.basename(os.path.normpath(os.path.abspath(self.path))))
            # Set UI
            self.pushbutton_getDataStart.setEnabled(False)
            self.url.setEnabled(False)
            self.r0.setEnabled(False)
            self.r1.setEnabled(False)
            self.checkbox_saveScreenshots.setEnabled(False)
            self.checkbox_saveImages.setEnabled(False)
            self.textedit_dataPath.setEnabled(False)
            self.pushbutton_loadData.setEnabled(False)
            self.treeWidget_vehicles.clear()
            self.log.clear()
            self.textedit_dataPath.setText(os.path.abspath(self.path))
            self.pushbutton_exportSelectedDataToExcel.setEnabled(False)
            self.progressBar.reset()
            # Clear dict(s)
            self.completedVehicles.clear()
            self.availableFields.clear()
            # Show data
            self.showData()
            # Reset count
            self.nMakeProcessed = 0
            # Save list of makes
            make_list = [name for name in os.listdir(self.path)]
            # Total number of makes
            self.nMakeTotal = len(make_list)
            # Process if make(s) exist
            if self.nMakeTotal > 0:
                # Set progress bar range
                self.progressBar.setRange(0, self.nMakeTotal)
                self.progressBar.setValue(0)
                # Iterate through directory structure
                for make in make_list:
                    vehicleList = list()
                    for model in os.listdir(self.path + "/" + make):
                        for year in os.listdir(self.path + "/" + make + "/" + model):
                            for name in os.listdir(self.path + "/" + make + "/" + model + "/" + year):
                                # Only process JSON files
                                if name.endswith(".json"):
                                    # Get ID
                                    m_obj = re.search(r"_(\d+)", name)
                                    if m_obj:
                                        id = m_obj.group(1) 
                                    else:
                                        raise Exception("Could not extract ID from JSON filename.")
                                    # Create object and add to list
                                    v = Vehicle(make, model, year, name, id, os.path.relpath(self.path) + "/" + make + "/" + model + "/" + year + "/" + name)
                                    vehicleList.append(v)
                    # Add to tree (per MAKE)
                    self.newMake(vehicleList)
            # Nothing to process
            if self.nMakeTotal == 0 or len(vehicleList) == 0:
                raise Exception("No data found.")
        except Exception, e:
            # Add to log
            self.addToLog("Could not load data: " + str(e))
            # Set UI
            self.pushbutton_getDataStart.setEnabled(True)
            self.pushbutton_getDataStop.setEnabled(False)
            self.url.setEnabled(True)
            self.r0.setEnabled(True)
            self.r1.setEnabled(True)
            self.checkbox_saveScreenshots.setEnabled(True)
            self.checkbox_saveImages.setEnabled(True)
            self.textedit_dataPath.setEnabled(True)
            self.pushbutton_loadData.setEnabled(True)
            self.pushbutton_exportSelectedDataToExcel.setEnabled(True)
            self.progressBar.reset()
            # Clear dict(s)
            self.completedVehicles.clear()
            self.availableFields.clear()
            # Show data
            self.showData()

    def saveFields(self, data_path):
        with open(data_path) as data_file:    
            data = json.load(data_file)
        # Iterate over field groups
        for key, value in data.iteritems():
            # Dictionary
            if isinstance(value, dict):
                if key not in self.availableFields:
                    self.availableFields[key] = set()
                    # Iterate over fields    
                    for k, v in value.iteritems():
                        self.availableFields[key].add(k)
                else:
                    # Iterate over fields    
                    for k, v in value.iteritems():
                        if k not in self.availableFields[key]:
                            self.availableFields[key].add(k) 
            # List
            else:
                if key not in self.availableFields:
                    self.availableFields[key] = set()
                    # Iterate over fields    
                    for i in value:
                        self.availableFields[key].add(i)
                else:
                    # Iterate over fields    
                    for i in value:
                        if i not in self.availableFields[key]:
                            self.availableFields[key].add(i)

    def showAllAvailableFields(self):
        data = self.availableFields
        for key, value in data.iteritems():
            item = self.addChild(self.treeWidget_data.invisibleRootItem(), 0, [key], None, True)
            # Add to tree widget
            for k in value:
                c = self.addChild(item, 0, [k], None, True)
        # Auto resize columns
        self.treeWidget_data.resizeColumnToContents(0);
        self.treeWidget_data.resizeColumnToContents(1);

    def showInFolder(self, item, col):
        file_path = item.data(col, QtCore.Qt.UserRole).toString()
        if os.path.isfile(file_path):
            dir_path = os.path.dirname(os.path.realpath(str(file_path)))
        else:
            dir_path = os.path.realpath(file_path)
        subprocess.Popen(r"explorer " + dir_path)

    def exportToExcel(self):
        # Do checks before exporting (checked tree widget items)
        # 1. Vehicles
        self.selectedVehicles = list()
        self.saveSelectedVehicles(self.treeWidget_vehicles.invisibleRootItem())
        if len(self.selectedVehicles) > 0:
            # 2. Parameters
            self.selectedParameters = list()
            self.saveSelectedParameters(self.treeWidget_data.invisibleRootItem())
            if len(self.selectedParameters) > 0:
                # Show confirmation message
                msg = "Vehicle(s) selected: " + str(len(self.selectedVehicles))
                msg += "\n" + "Parameter(s) selected: " + str(len(self.selectedParameters))
                msg += "\n\n" +  "Do you wish to proceed?"
                reply = QtGui.QMessageBox.question(self, 'Export to Excel', msg, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                if reply == QtGui.QMessageBox.Yes:
                    # Do Excel export
                    self.doExcelExport()
            else:
                QtGui.QMessageBox.warning(self, "Error", "Cannot export to Excel: No parameters selected")
        else:
            QtGui.QMessageBox.warning(self, "Error", "Cannot export to Excel: No vehicles selected")

    def saveSelectedVehicles(self, item):
        # Check if item is vehicle (has data which points to a file) and is checked
        if item.data(0, QtCore.Qt.UserRole) != None and os.path.isfile(item.data(0, QtCore.Qt.UserRole).toString()) and item.checkState(0) == QtCore.Qt.Checked:
            self.selectedVehicles.append(item)
        # Check children (recursively)
        for i in range(0, item.childCount()):
            self.saveSelectedVehicles(item.child(i))

    def saveSelectedParameters(self, item):
        # Check if item is parameter (has no children) and is checked
        if item.childCount()==0 and item.checkState(0) == QtCore.Qt.Checked:
            self.selectedParameters.append(item)
        # Check children (recursively)
        for i in range(0, item.childCount()):
            self.saveSelectedParameters(item.child(i))

    def doExcelExport(self):
        # Get save file path from user
        self.filename = QtGui.QFileDialog.getSaveFileName(self, "Save Workbook", "RedBookData" + "_" + str(self.timestamp) + "_" + str(len(self.selectedVehicles)) + "vehicles.xlsx", ".xlsx")
        # Create workbook
        if self.filename != "":
            # Set progress bar range
            self.nVehiclesExported = 0
            self.progressBar_export.setRange(0, len(self.selectedVehicles))
            self.progressBar_export.setValue(0)
            # Set UI
            self.pushbutton_getDataStart.setEnabled(False)
            self.pushbutton_exportSelectedDataToExcel.setEnabled(False)
            self.pushbutton_loadData.setEnabled(False)
            self.log.clear()
            # Create and start thread
            self.exportThread = ExportThread(self.filename, self.timestamp, self.selectedVehicles, self.selectedParameters)
            self.exportThread.finished.connect(self.excelExportDone, QtCore.Qt.QueuedConnection)
            self.exportThread.vehicleDone.connect(self.excelExportVehicle, QtCore.Qt.QueuedConnection)
            self.exportThread.start()

    def excelExportDone(self):
        # Set UI
        self.pushbutton_getDataStart.setEnabled(True)
        self.pushbutton_exportSelectedDataToExcel.setEnabled(True)
        self.pushbutton_loadData.setEnabled(True)
        self.addToLog("Finished! Excel workbook saved to: " + str(self.filename))
        # Set thread to None
        self.exportThread = None

    def excelExportVehicle(self, vehicle):
        self.nVehiclesExported += 1
        self.progressBar_export.setValue(self.nVehiclesExported)
        if self.nVehiclesExported == len(self.selectedVehicles):
            self.addToLog("Saving Excel file. Please wait... (this may take several minutes)")
        else:
            self.addToLog("Processed: " + vehicle)

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
