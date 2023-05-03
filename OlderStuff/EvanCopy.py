import os
import re
import pdb
import csv
import time
import glob
import ntpath
import random
import logging
import pandas as pd
from pathlib import Path
from pudb import set_trace
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementClickInterceptedException
print('What is the desired program speed? (enter an time, in seconds)')
while True:
    try:
        speed = int(input())
        break
    except:
        print("Wrong type of inpt, please enter a integer to be used at critical sleep points")
print("Program speed of " + str(speed) + " seconds at major sleep points")
# TODO: Add paths
desktop_directory = r'C:\Users\WillKnight\Desktop'
download_directory = r'C:\Users\WillKnight\Downloads'
csv_path = r'C:\Users\WillKnight\Dropbox\january\newcsv.csv'
skipped_csv_path = r'C:\Users\WillKnight\Dropbox\january\skippedcsv.csv'
log_path = r'C:\Users\WillKnight\Dropbox\january\2012Part3.log'
#skipped = r:'C\Users\WillKnight\Dropbox\january\skipped.csv'
#For Mac:
#chromedriver = "/Users/will_knight/Desktop/chromedriver"
#For Windows:
#chromedriver = r"c:\Users\WillKnight\Desktop\chromedriver"
chromedriver = os.path.join(desktop_directory, 'chromedriver')
os.environ["webdriver.chrome.driver"] = chromedriver
chrome_options = Options() #old
prefs = {'profile.default_content_setting_values.automatic_downloads': 1}
#chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_experimental_option("prefs", prefs)
#options.addArguments("--disable-gpu")
#options.add_arguments("--headless --disable-gpu") #new
# TODO: Add multiple drivers to restart if times tried is greater than 3
driver = webdriver.Chrome(chromedriver, chrome_options=chrome_options)
df = pd.read_csv("updated_1500.csv")
cusipsList = df.cusip6.to_list()
cusipList = df.cusip.to_list()
conmList = df.conm.to_list()
ticsList = df.tic.to_list()
#sinList = df.Sin.to_list() Bring up to Brandon, No idea why in the world this dowesn't work
gvkeyList = df.gvkey.to_list()
is_sp500List = df.is_sp500.to_list()
is_sp1500List = df.is_sp1500.to_list()
report_number = 0
current_cusip = cusipList[0]
current_cusip6 = cusipsList[0]
current_conm = conmList[0]
current_tic = ticsList[0]
current_gvkey = gvkeyList[0]
current_is_sp500 = is_sp500List[0]
current_is_sp1500 = is_sp1500List[0]
currentCusip = str(cusipsList[0])
currentTic = str(ticsList[0])
fieldnames = ['downloadNumber', 'reportNumber', 'reportID', 'reportTitle', 'Date', 'Author', 'Path', 'ExcelPath', 'cusip6', 'cusip', 'conm', 'tic', 'gvkey',  'timeStamp', 'is_sp500', 'is_sp1500', 'download_status']
#For Mac:
#logging.basicConfig(filename='betaTest.log', encoding='utf-8', level=logging.DEBUG)
#For Windows:
#logFile = os.path.join(desktop_directory, "Dropbox", "betaTest.log")
#logging.basicConfig(filename="logtest.log", level=logging.INFO)
logging.basicConfig(filename=log_path, level=logging.INFO)
years = ['2012','2013','2014','2015','2016','2017','2018','2019','2020','2021']
months = ['1-Jan','2-Feb','3-Mar','4-Apr','5-May','6-Jun','7-Jul','8-Aug','9-Sep','10-Oct','11-Nov','12-Dec']
#EVAN YANG- commented out accounts below
accounts = ['gbsstudent1@emory.edu', 'gbsstudent2@emory.edu','gbsstudent3@emory.edu', 'gbsstudent4@emory.edu', 'gbsstudent6@emory.edu', 'gbsstudent7@emory.edu', 'gbsstudent8@emory.edu', 'gbsstudent10@emory.edu','gbsstudent11@emory.edu', 'gbsstudent12@emory.edu', 'gbsstudent13@emory.edu','gbsstudent14@emory.edu', 'gbsstudent15@emory.edu']

#accounts=['gbsstudent6@emory.edu']

#with open('C:/Users/WillKnight/Dropbox/betaTest.csv', 'a', newline='') as csvfile:
#with open(csv_path, 'a', newline='') as csvfile:
#    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
#    writer.writeheader()
logging.info("\n\nPROGRAM EXECUTION: " + time.asctime() + "")

def main():
    global company,year,month
    global accounts
    #Evan- start Parameters
    company = 919 #Evan- change company here, companyChange
    year = 1 #Evan- Changed from 0 to 1 to do 2013
    month = 0
    report_number = 0
    global current_account, companyErrorNumber
    current_account = change_account(0)
    login(current_account)
    if checkparameters() == False:
        logging.info("Account number" + str(accounts[current_account]) + " has incorrect ordering of parameters")
        login(change_account(current_account))
    while year < 10:
        global currentYear
        currentYear = years[year]
        while company < 2033: #Evan- changed from 2035 to 2033
            global current_cusip, current_cusip6, current_conm, current_tic, current_gvkey, current_is_sp500, current_is_sp1500, currentCusip, currentTic, speed
            current_cusip = cusipList[company]
            current_cusip6 = cusipsList[company]
            current_conm = conmList[company]
            current_tic = ticsList[company]
            current_gvkey = gvkeyList[company]
            current_is_sp500 = is_sp500List[company]
            current_is_sp1500 = is_sp1500List[company]
            currentCusip = str(cusipsList[company])
            currentTic = str(ticsList[company])
            logging.info("On Company number: " + str(company) + " which is: " + currentTic)
            companyErrorNumber = 0
            while True:
                try:
                    time.sleep(speed)
                    clear()
                    start = startYear(currentYear)
                    end = endYear(currentYear)
                    year_files_in_search = search_and_download(currentCusip, currentYear, currentTic, start, end)
                    break
                except:
                    logging.error("Crashed on Search:" + currentTic + " - " + currentCusip + " - " + currentYear)
                    companyErrorNumber = companyErrorNumber + 1
                    cleanup(currentCusip, currentYear, currentTic)
                    newDriver()
                    login(change_account(current_account))
                    if(companyErrorNumber > 2):
                        skip_company()
                        companyErrorNumber = 0
                        break
            companyErrorNumber = 0
            if year_files_in_search == -5: #If in the yearly search theres no reports available, logs none and skips the year
                logging.info("No reports available to be downloaded for company: " + currentCusip +", " + currentTic + " in year: " + str(currentYear))
            elif 101 > year_files_in_search > 0:
                place_in_directory(currentCusip, currentYear, currentTic, start, end)
            elif year_files_in_search > 100: #If there's more than 100, then it switches to quarterly search and downloads for that Year
                logging.info("Switching to quarterly downloading for " + currentCusip + ", " + currentTic + " in year: " + str(currentYear) + ". There were " + str(year_files_in_search) + " reports.")
                quarter = 0
                while quarter < 4:
                    start = quarterStart(quarter, currentYear)
                    end = quarterEnd(quarter, currentYear)
                    companyErrorNumber = 0
                    while True:
                        try:
                            clear()
                            quarter_files_in_search = search_and_download(currentCusip, currentYear, currentTic, start, end)
                            place_in_directory(currentCusip, currentYear, currentTic, start, end)
                            break
                        except:
                            companyErrorNumber = companyErrorNumber + 1
                            logging.error("Crashed on Search:" + currentTic + " - " + currentCusip + " - " + start + " through " + end)
                            cleanup(currentCusip, currentYear, currentTic)
                            newDriver()
                            login(change_account(current_account))
                            if(companyErrorNumber > 2):
                                skip_company()
                                companyErrorNumber = 0
                                break
                    companyErrorNumber = 0
                    if(quarter_files_in_search > 100):
                        logging.info("Switching to Monthly search for the quarter: " + str(quarter))
                        if(quarter == 0):
                            month = 0
                        elif(quarter == 1):
                            month = 3
                        elif(quarter == 2):
                            month = 6
                        elif(quarter == 3):
                            month = 9
                        plusThree = month + 3
                        while month < plusThree :
                            start = monthStart(month, currentYear)
                            end = monthEnd(month, currentYear)
                            print("start: " + str(start))
                            print("end: " + str(end))
                            companyErrorNumber = 0
                            while True:
                                try:
                                    clear()
                                    month_files_in_search = search_and_download(currentCusip, currentYear, currentTic, start, end)
                                    break
                                except:
                                    companyErrorNumber = companyErrorNumber + 1
                                    logging.error("Crashed on Search:" + currentTic + " - " + currentCusip + " - " + start + " through " + end)
                                    cleanup(currentCusip, currentYear, currentTic)
                                    newDriver()
                                    login(change_account(current_account))
                                    if(companyErrorNumber > 2):
                                        skip_company()
                                        companyErrorNumber = 0
                                        break
                            while True:
                                try:
                                    place_in_directory(currentCusip, currentYear, currentTic, start, end)
                                    break
                                except:
                                    logging.error("Crashed placing file on Search:" + currentTic + " - " + currentCusip + " - " + start + " through " + end)
                                    cleanup(currentCusip, currentYear, currentTic)
                                    newDriver()
                                    login(change_account(current_account))
                            month = month + 1
                    while True:
                        companyErrorNumber = 0
                        try:
                            place_in_directory(currentCusip, currentYear, currentTic, start, end)
                            break
                        except:
                            companyErrorNumber = companyErrorNumber + 1
                            logging.error("Crashed placing file on Search:" + currentTic + " - " + currentCusip + " - " + start + " through " + end)
                            cleanup(currentCusip, currentYear, currentTic)
                            newDriver()
                            login(change_account(current_account))
                    logging.info("Finished searching/downloading for quarter " + str(quarter) + " for: " + currentTic + ", " + currentCusip + " in " + str(currentYear))
                    quarter = quarter + 1
            elif year_files_in_search == -1:
                logging.error("Error fetching data on for company: " + currentCusip + " in year: " + str(currentYear))
            logging.info("Finished searching and downloading for " + currentCusip + " in year " + str(currentYear))
            company = company + 1

            #############################
            if(company % 40 == 0): #Changed from 25 to 40- Evan Yang
                logging.info("Finished with 40 Searches, Switching Accounts and Drivers to Balance Load")
                driver.quit() # quits the new driver
                driver = webdriver.Chrome(chromedriver, chrome_options=chrome_options) #makes a new driver
                login(change_account(current_account))
                logging.info("Accoount Switch Successful")
            ###########################
            
            logging.info("Moving to the next Company, and the time is currently: " + time.asctime())
        year = year + 1
        logging.info("Finished with the Year: " + str(currentYear) )
        time.sleep(speed)

def login(accountNumber):
    login_error = 0
    global driver
    while True:
        try:
            driver.get('http://eikon.thomsonreuters.com/index.html')
            time.sleep(.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div/div[2]/a"))).click()
            break
        except TimeoutException:
            if(login_error > 2):
                raise Exception ("Error Loading Sign-In Page, throwing error")
            logging.warning("Error loading database, likely a connectivity issue")
            login_error = login_error + 1
    while True:
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div/div[2]/form/table/tbody/tr[1]/td[2]/div")))
            break
        except TimeoutException:
            try:
                driver.switch_to_default_content()
                WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
                WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'helios_button'))).click()
                WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('heliosMenuFrame')))
                WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
                WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/eikon-app-window/div/div[9]/div/div/eikon-app-body/div/div[1]/div/div[1]/div[2]/div/div[7]'))).click()
                driver.switch_to_default_content()
                WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'AAA-AS-SI5-SE003'))).click()
                #WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'AAA-AS-SI5-SE003'))).click()
                break
            except:
                raise Exception ("Error Loading Sign-In Page, throwing error")
            break
    time.sleep(.2)
    usernameBox = driver.find_element_by_xpath('//*[@id="AAA-AS-SI1-SE003"]')
    currentUser = accounts[accountNumber] #TypeError: list indices must be integers or slices, not function?
    usernameBox.send_keys(currentUser)
    time.sleep(.2)
    passwordBox = driver.find_element_by_xpath('//*[@id="AAA-AS-SI1-SE006"]')
    passwordBox.send_keys('Balance5$')
    driver.find_element_by_id('AAA-AS-SI1-SE014').click()
    try:
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID,'AAA-AS-SI2-SE009'))).click()
    except TimeoutException:
        logging.debug("Error:  Account Already in Use Popup, breaking instead")
    check = driver.page_source
    error_logging = check.find("Incorrect User")
    if error_logging > 0:
        login(change_account(accountNumber))
        return
    time.sleep(.5)
    driver.get('https://amers1.apps.cp.thomsonreuters.com/web/apps/research#/') #navigates to advanced research page automatically rather than through navigation
    logging.info("Fully logged in on account: " + accounts[accountNumber] )
def logout():
    #accounts = ['gbsstudent1@emory.edu', 'gbsstudent2@emory.edu', 'gbsstudent4@emory.edu', 'gbsstudent6@emory.edu', 'gbsstudent7@emory.edu', 'gbsstudent8@emory.edu', 'gbsstudent10@emory.edu', 'gbsstudent12@emory.edu', 'gbsstudent14@emory.edu']
    global current_account;
    global accounts;
    logout_error = 0
    while True:
        try: #This is the orioginal method of logging out
            driver.switch_to_default_content()
            WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'helios_button'))).click()
            driver.switch_to_default_content()
            WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('heliosMenuFrame')))
            WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/eikon-app-window/div/div[9]/div/div/eikon-app-body/div/div[1]/div/div[1]/div[2]/div/div[7]'))).click()
            driver.switch_to_default_content()
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'AAA-AS-SI5-SE003'))).click()
            break
        except: #2nd method of logging out whcih is on the right hand side of the screen
            #pdb.set_trace()
            logging.info('Cannot press the Eikon button to have the eikon menu drop down')
            raise Exception("Error pressing Eikon button during logout")
            #windows = driver.window_handles
            #if(len(windows) > 1):
            #    driver.switch_to_window(windows[1])
            #    driver.close()
            #    windows = driver.window_handles
            #    driver.switch_to_window(windows[0])
            #if(logout_error > 2):
            #    newDriver()
            #    logout_error = 0
    logging.info("Logged out of account: " + accounts[current_account])
def change_account(current_Number):
    global accounts
    while True:
        length_of_accounts = len(accounts) # Will automatically adjust to number of accounts functional
        newNum = random.randint(0, length_of_accounts - 1)
        if newNum != current_Number:
            break
    return newNum
def clear():
    clear_error = 0
    while True:
        try:
            driver.switch_to_default_content()
            WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'helios_button')))
            break
        except TimeoutException:
            clear_error = clear_error + 1
            if(clear_error > 2):
                #pdb.set_trace()
                raise Exception("Error switching frames in Clear Method")
    if driver.current_url != 'https://amers1.apps.cp.thomsonreuters.com/web/apps/research#/':
        driver.get('https://amers1.apps.cp.thomsonreuters.com/web/apps/research#/')
    clear_error = 0
    while True:
        try:
            driver.switch_to_default_content()
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('contentframe')))
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            break
        except TimeoutException:
            clear_error = clear_error + 1
            if(clear_error > 2):

                raise Exception("Error switching frames on advanced search page")
    clear_error = 0
    while True:
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, 'clearButton'))).click()
            break
        except TimeoutException:
            clear_error = clear_error + 1
            if(clear_error > 2):
                #pdb.set_trace()
                raise Exception("Error finding the clear button")
    times_tried = 0
    while True:
        clicked = False
        for i in range(5, 20):
            try:
                okayButton =  '/html/body/div[' + str(i) + ']/div/div/div[3]/button[1]'
                WebDriverWait(driver, .3).until(EC.presence_of_element_located((By.XPATH, okayButton))).click()
                clicked = True
                break
            except TimeoutException:
                logging.debug("okayButton: " + okayButton + " didn't work, trying alternate")
        if clicked:
            break
        times_tried = times_tried + 1
        if times_tried > 3:
            driver.find_element_by_id('clearButton').click()
def checkDuplicate(folderName, fileName):
    for name in os.listdir(folderName):
        if name == fileName:
            #set_trace()
            downloadedFile = os.path.join( r'C:\Users\WillKnight\Downloads', name)
            os.remove(downloadedFile)
            logging.info("Duplicate found for file " + name + ' in the folder: ' + folderName)
            return True
    return False
def checkparameters():
        global current_account
        driver.get('https://amers1.apps.cp.thomsonreuters.com/web/apps/research#/')
        tried = 0
        time.sleep(1)
        while tried < 3:
            try:
                driver.switch_to_default_content
                WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
                WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('contentframe')))
                WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
                break
            except TimeoutException:
                tried = tried + 1
        times_tried = 0
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'clearButton')))
        except TimeoutException:
            if(times_tried > 2):
                return False
            times_tried = times_tried + 1
            print("Likely error when looking for clear")
            driver.refresh()
            time.sleep(5)
            driver.get('https://amers1.apps.cp.thomsonreuters.com/web/apps/research#/')
        #try:
        #    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/header/div[3]/div[2]/span/button'))).click()
        #except TimeoutException:
        #    logging.debug("cannot press settings")
        #try:
        #    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/header/div[3]/div[2]/span/ul/li[8]/a'))).click()
        #except TimeoutException:
        #    logging.debug("cannot select dowloading options")
        #try:
        #    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[3]/div/div[2]/div[1]/ul/li[1]'))).click()
        #except TimeoutException:
        #    logging.debug("cannot select basic download preferences")
        #try:
        #    time.sleep(1)
        #    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="prefs-panel-right"]/div/div[2]/div[2]/input[1]'))).click()
        #except TimeoutException:
        #    print("cannot close download preference tab")
        page = driver.page_source
        dateRange = page.find("Date Range")
        keyword = page.find("Keyword Search")
        companys = page.find("Companies")
        portfolio = page.find("Portfolios")
        reportType = page.find("Report Types")
        subject = page.find("Subjects/Topics")
        pageCount = page.find("Page Count")
        analysts = page.find("Analysts")
        fileType = page.find("File Types")
        country = page.find("Countries/Regions")
        docID = page.find("Document ID")
        quick = page.find("Quick Search")
        industry = page.find("Industry")
        if 0 <= dateRange <= keyword <= companys <= portfolio <= reportType <= subject <= pageCount <= analysts <= fileType <= country <= docID <= quick <= industry:
            logging.info("Parameters are in correct order for account: " + str(accounts[current_account]))
            return True
        return False
def search_and_download(Cusip, Year, Tic, startDate, endDate):
    parameters = driver.find_elements_by_class_name('select2-choice') #Parameters holds array of dropdown menu elements to be referenced
    setDateRange = parameters[1]
    escape_error = 0
    while True:
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/date-criteria/div[2]/div/a'))).click()
            time.sleep(1)
            for i in range(0, 12):
                    setDateRange.send_keys(Keys.ARROW_DOWN)
            setDateRange.send_keys(Keys.ENTER)
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if (escape_error > 2):
                logging.error("Cannot select Custom Date Range")
                raise Exception("Cannot select Custom Date Range")
    #This portion enters the first and second dates, however, they must be in the correct format: 01-Jul-2018 ; 01-Mon-Year
    dateRange1 = startDate
    dateRange2 = endDate
    errorNum = 0
    tab = ActionChains(driver)
    tab.send_keys(Keys.TAB)
    escape_error = 0
    while True: #This enters the first date range into the dateFrom parameter
            try:
                    time.sleep(1)
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//input[starts-with(@id, "dateFrom-")]')))
                    first_date = driver.find_element_by_xpath('//input[starts-with(@id, "dateFrom-")]')
                    driver.execute_script('arguments[0].value = "";', first_date)
                    driver.execute_script("arguments[0].click();", first_date)
                    first_date.send_keys(dateRange1)
                    first_date.send_keys(Keys.ESCAPE)
                    break
            except TimeoutException:
                    logging.debug("Switching to alternative for entering the first date")
                    tab.perform()
                    errorNum = errorNum + 1
                    if errorNum > 2:
                        try:
                            dateNumOne = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/date-criteria/div[2]/div').click()
                        #/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/date-criteria/div[2]/div
                        #dateNumOne = driver.find_element_by_id('s2id_namedRange-4820').click()
                            down = ActionChains(driver)
                            down.send_keys(Keys.ARROW_DOWN)
                            for i in range (0, 10):
                                down.perform()
                            tab = ActionChains(driver)
                            tab.send_keys(Keys.TAB)
                            tab.perform()
                            dateNumOne = ActionChains(driver)
                            dateNumOne.send_keys(dateRange1)
                            dateNumOne.perform()
                            escape = ActionChains(driver)
                            escape.send_keys(Keys.ESCAPE)
                            escape.perform()
                            break
                        except:
                            logging.error("Error entering the first date")
                            raise Exception("Error clicking on the DateFrom Box")
    escape_error = 0
    while True: #This enters the second date range into the dateTo parameter
            try:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//input[starts-with(@id, "dateTo-")]')))
                    second_date = driver.find_element_by_xpath('//input[starts-with(@id, "dateTo-")]')
                    driver.execute_script('arguments[0].value = "";', second_date)
                    driver.execute_script("arguments[0].click();", second_date)
                    second_date.send_keys(dateRange2)
                    second_date.send_keys(Keys.ESCAPE)
                    break
            except TimeoutException:
                    escape_error = escape_error + 1
                    if(escape_error > 2):
                        raise Exception("Error: issue while entering Date-Range To")
                    logging.error("Error: issue while entering Date-Range To")
    escape_error = 0
    while True: #This portion enters the Cusip into the search
        try:
            criteria_elements = driver.find_elements_by_xpath('//input[starts-with(@id, "criteria-")]')
            Companies_Box = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/ul/li[2]/div/company-criteria/div[2]/div[1]/etk-type-ahead-control/input')
            Companies_Box.clear()
            Companies_Box.click()
            Companies_Box.send_keys(Cusip)
            break
        except:
            escape_error = escape_error + 1
            if(escape_error > 2):
                logging.error("Error entering the second date, throwing an error to get a new webpage")
                raise Exception("Error entering second date")

    escape_error = 0
    while True: #This selects the top match in the dropdown menu after searching for the Cusip6
                    try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/ul/li[1]/div/div'))).click()
                            break
                    except TimeoutException:
                        escape_error = escape_error + 1
                        if escape_error > 2:
                            logging.error("Skipping Company " + str(company) + ": Could not find Ticker:" + str(currentTic) + " in year " + str(year))
                            raise Exception("Error finding ticker after putting in: " + str(currentTic) + " in Year: " + str(Year) )
                            #TODO: Add a Write to Skipped.csv here

                                #with open(skipped, 'a', newline='') as csvfile:
                                #    writer = csv.DictWriter(csvfile, fieldnames=fieldnames_skipped)
                                #    no_ticker
                                #    writer.writerow({'Error':no_ticker, 'Year':})
                                #return False
    escape_error =  0
    while True: #This presses done on the dropdown so the script can enter more parameters
                    try:
                            WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/button[4]'))).click()
                            break
                    except TimeoutException:
                        escape_error =  escape_error + 1
                        if (escape_error > 2):
                            logging.error("Isssues clicking on the 'Done' ")
                            raise Exception("Issues clicking on the 'Done' ")
    escape_error =  0
    while True:     #This code block limits the Search Providers to not allow independent or unverified publishers
                    try:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/ul/li[3]/div/contributor-criteria/div[2]/div[2]/etk-dropdown-tree/div/input'))).click()
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/ul/li[2]/span/label/span'))).click()
                        time.sleep(1) #For some reason this little sleep allows the button to be clicked, even with the element_to_be_clickable attribute
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/ul/li[3]/span/label/span'))).click()
                        driver.find_element_by_xpath('/html/body/div[4]/div/div[3]/button[4]').click()
                        break
                    except TimeoutException:
                        escape_error =  escape_error + 1
                        if (escape_error > 2):
                            logging.error("Cannot limit the Research Providers")
                            raise Exception("Cannot limit the Research Providers")
    escape_error =  0
    while True: #This code block selects the primary company box to ensure the company searched for is the primary company of the report
                    try:
                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/ul/li[4]/div/div[2]/div[2]/label/span'))).click()
                            break
                    except TimeoutException:
                        escape_error =  escape_error + 1
                        if (escape_error > 2):
                            logging.error("Cannot select 'Primary Company' box")
                            raise Exception("Cannot select 'Primary Company' box")
    escape_error =  0
    while True:     #This code block scrolls down to have the Nation parameter be in view
                    try:
                        down = ActionChains(driver)
                        down.send_keys(Keys.ARROW_DOWN)
                        for i in range(0, 12):
                            down.perform()
                        break
                    except TimeoutException:
                        escape_error =  escape_error + 1
                        if (escape_error > 2):
                            logging.error("Cannot scroll down to access Nation and FileType paramaters")
                            raise Exception("Cannot scroll down to access Nation and FileType paramaters")
    escape_error =  0
    while True:     #This code block limits the file Types to be strictly PDFs
                    try:
                        #set_trace()
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/ul/li[10]/div/file-types-criteria/div[2]/etk-dropdown-tree/div/input'))).click()
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[9]/div/div[1]/ul/li[2]/span/label/span'))).click()
                        driver.find_element_by_xpath('/html/body/div[9]/div/div[3]/button[3]').click()
                        break
                    except TimeoutException:
                        escape_error =  escape_error + 1
                        if (escape_error > 2):
                            logging.error("Cannot limit the file Type to be PDF only")
                            raise Exception("Cannot limit the file Type to be PDF only")
                    except ElementClickInterceptedException:
                        driver.find_element_by_xpath('/html/body/div[10]/div/div[1]/ul/li[4]/span/label/span').click()
                        driver.find_element_by_xpath('/html/body/div[10]/div/div[3]/button[4]').click()
                        break
    escape_error =  0
    while True:     #This code block limits the publisher's Nations to be the "Americas" Only (No United States)
                    try:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div[2]/div[1]/div/div[1]/form/ul/li[11]/div/region-criteria/div[2]/etk-dropdown-tree/div/input'))).click()
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div/div[1]/ul/li[4]/span/label/span'))).click()
                        driver.find_element_by_xpath('/html/body/div[10]/div/div[3]/button[4]').click()
                        break
                    except TimeoutException:
                        escape_error =  escape_error + 1
                        if (escape_error > 2):
                            logging.error("Cannot Select the United States as Locations")
                            raise Exception("Cannot Select the United States as Locations")
                    except ElementClickInterceptedException:
                        driver.find_element_by_xpath('/html/body/div[9]/div/div[3]/button[3]').click()
                        break

    Companies_Box.send_keys(Keys.ENTER)
    number = 0 #This portion of code is after the search, it looks for the search results to pop up through parsing the page file for "Displaying"
    while True:
        again = False
        for i in range (0, 20): # TODO: Find the optimal time for this sleep; can be 5-20 roughly
            time.sleep(1)
            page = driver.page_source
            yesReports = page.find("Displaying ")
            if( yesReports > -1): #If it knows reports are there, it grabs the number of reports via RegEX to be used later during report downloading
                again = True
                pattern = re.compile(r"(Displaying )(\d*)( of )(\d*)")
                for match in pattern.finditer(page):
                    number = int(match.group(4))
                break
        if again:
            break
        noReports = page.find('No data available') #If after a bit there is no reports loading, it checks for Error fetching data and No data; both of which skip this search
        errorLoading = page.find('Error fetching data')
        if( errorLoading > -1):
            logging.error("Error Fetching Data!")
            return -1
        if( noReports > -1):
            logging.info("No Data Available")
            return -5

    if number > 100: #This stops the below from occuring
        return number

    escape_error = 0
    while True: #This portion clicks on the 'select all' button to get every report on the search page
        try:
            selectAll = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'trgrid_cbcplugin_cb0')))
            driver.execute_script("arguments[0].click();", selectAll)
            break
        except TimeoutException:
            logging.error("Switching to xpath method of getting every report")
        try:
            selectAll = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/div[2]/div/div[1]/div[1]/div/div/div[1]/div[1]/div[1]/svg/svg/div[1]/div/div/input')))
            driver.execute_script("arguments[0].click();", selectAll)
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if(escape_error > 2):
                raise Exception("Issues selecting all the reports on the page")
            logging.warning("Issues selecting all the reports on the page")
    escape_error - 0
    while True: #This selects the Dropdown menu
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/button[2]'))).click()
            #WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/button[2]'))).click()
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Cannot press the dropdown menu to export to Excel")
            logging.error("Cannot press the dropdown menu to export to Excel")
    escape_error = 0
    while True: #This exports the report's information to an excel file and downloads it
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/ul/li[2]/excel-export/a'))).click()
            #WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/ul/li[2]/excel-export/a'))).click()
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Error pressing the 'Export to Excel' button")
            logging.error("Error pressing the 'Export to Excel' button")

    escape_error = 0
    while True: #This presses the dropdown twice to refresh it so that the PDFs can be downloaded
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/button[2]'))).click()
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/button[2]'))).click()
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Cannot press the dropdown menu to export to Excel 2")
            logging.error("Cannot press the dropdown menu to export to Excel 2")

    escape_error = 0
    mainPage = driver.current_window_handle #Saves the current page to be returned to later
    #set_trace()
    while True: #Presses on download/save as PDFs and the new window pops up
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/ul/li[1]/a'))).click()
            #WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/div/header/div[2]/div[6]/div/ul/li[1]/a'))).click()
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Error getting PDFs to Download into popup window")
            logging.debug("Error getting PDFs to Download into popup window")
    for i in range(0, 20):
        time.sleep(1)
        windows = driver.window_handles
        if len(windows) == 2:
            break
    time.sleep(2)
    escape_error = 0
    while True:
        try:
            windows = driver.window_handles
            driver.switch_to_window(windows[1])
            driver.switch_to_default_content()
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('contentframe')))
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Error switching frames in new window")
            logging.debug("Error switching frames in new window")

    escape_error = 0
    while True:
        try:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div[1]/div[4]/div/div[1]/div[1]/table/thead/tr/td[4]/a/span[1]')))
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Page Cannot load with all the PDFs ready to be downloaded")
            logging.debug("Page Cannot Load with all the PDFs ready to be downloaded")
    #set_trace()
    escape_error = 0
    while True:
        try:
            time.sleep(.2)
            downloaded_files = len(os.listdir(r'C:\Users\WillKnight\Downloads')) #This gets the number of files in downloads to be referenced later
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div[1]/div[3]/div[2]/button[2]'))).click()
            break
        except TimeoutException:
            escape_error = escape_error + 1
            if( escape_error > 2):
                raise Exception("Error loading the new page with multiple PDFs")
            logging.debug("Loading new page with multiple PDFs")
        except ElementClickInterceptedException:
            try:
                time.sleep(3)
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div[1]/div[3]/div[2]/button[2]/span'))).click()
                break
            except TimeoutException:
                logging.debug("Error pressing pdfs to download")
            except ElementClickInterceptedException:
                driver.find_element_by_xpath('/html/body/div/div/div/div[1]/div[3]/div[2]/button[2]').click()
                break
    #By now, the files will begin to download, could be a long ass process
    time.sleep(5)
    for i in range(0, 20):
        time.sleep(1)
        windows = driver.window_handles
        if len(windows) == 2:
            break
    if (len(driver.window_handles) != 2):
        raise Exception("Popup after searching not coming up. Raising exception")

    driver.switch_to_window(windows[1])
    driver.switch_to_default_content()
    WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
    WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('contentframe')))
    WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
    driver.execute_script("document.body.style.zoom='20%'")
    # TODO: Figure out correct sizing/zoom to fit at most 100 reports
    driver.set_window_size(1200, 1000)
    driver.execute_script("$('#content').css('zoom', 5);")
    # TODO: use page source to read saved vs error
    #This is the downloading portion of the pdfs, these two helper varaibles give the ability to break the while loop below
    goal_files = downloaded_files + number
    startTime = time.time()
    while True: #This portion checks whether or not every single file has been downloaded completely or if it's been incessantly long and something went wrong
        number_of_files = len(os.listdir(r'C:\Users\WillKnight\Downloads'))
        page = driver.page_source
        status = page.count('Saved')
        errornum = page.count('Error')
        if errornum > 0:
            #set_trace()
            goal_files = downloaded_files + number - errornum
            logging.error("Incorrect number of files downloaded during search, redoing this search")
            raise Exception("Error while downloading files")
        if number_of_files == goal_files:
            time.sleep(1) #This allows the final file to completely download, may need a way to verify quicker than just a wait
            print(time.time() - startTime)
            break
        elif time.time() - startTime > 420: #To Download: 107. 128, 216,226
            logging.error("Issues downloading correct number of PDFs")
            logging.error("Should have downloaded: " + str(number) + "files, but only downloaded " + str(number_of_files - downloaded_files))
            break
    #yearFolder = '/Users/will_knight/Desktop/BetaTest/' + Cusip + '_' + Tic + '/' + Year + '/Advanced Research Search.xlsx'
    time.sleep(1) #This portion closes the download window popup and returns to the original page
    windows = driver.window_handles
    driver.switch_to_window(windows[1])
    driver.switch_to_default_content()
    WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
    WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('contentframe')))
    WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it(('AppFrame')))
    page = driver.page_source
    status = page.count('Saved')
    if status == number:
        #set_trace()
        driver.close()
        driver.switch_to_window(mainPage)
        return number
    else:
        logging.error("Issue downloading correct number of files; missing " + str(number_of_files - downloaded_files) + " files from downloads")
        return number
def place_in_directory(Cusip, Year, Tic, startDate, endDate):
    global desktop_directory, download_directory
    cusipYear = Cusip + "_" + Tic
    yearString = str(Year)
    excelName = Cusip + '_' + Tic + '_' + startDate + '-' + endDate + '.xlsx'
    #excelName = Cusip + '_' + Tic + '_01-01-' + Year + '_' + '12-31-' + Year + '_Y.xlsx'
    helper = excel_folder = os.path.join(desktop_directory, "BetaTest", cusipYear, yearString[2:4])
    excel_folder = os.path.join(desktop_directory, "BetaTest", cusipYear, yearString[2:4], excelName)
    local_report_number = 0
    for i in range(0, 12):
        start_of_file = str(Year) + '-' + convert3(months[i])
        for name in os.listdir(r'C:\Users\WillKnight\Downloads'):
            if(name.startswith("Advanced")):
                researchSearch = os.path.join(download_directory, name)
                if os.path.exists(helper) == False:
                    os.makedirs(helper)
                    time.sleep(.01)
                try:
                    time.sleep(.01)
                    os.rename(researchSearch, excel_folder)
                except:
                    logging.debug("No advanced Search excel file exists to be renamed")
            if(name.startswith(start_of_file)):
                time.sleep(.01)
                fileDestination = os.path.join(desktop_directory, "BetaTest", cusipYear, yearString[2:4], convert7(months[i]))
                if os.path.exists(fileDestination) == False:
                    os.makedirs(fileDestination)
                    time.sleep(.01)
                currently_here = os.path.join(download_directory, name)
                newName = os.path.join(fileDestination, name)
                if checkDuplicate(fileDestination, name):
                     return
                time.sleep(.01)
                os.rename(currently_here, newName)
                split = re.search(r"(\d{4}-\d{2}-\d{2})-(.*)-(.*)-(.*)-(\d*)", name)
                global report_number, current_cusip6, current_cusip, current_conm, current_tic, current_gvkey, current_is_sp500, current_is_sp1500
                report_number = report_number + 1
                local_report_number = local_report_number + 1
                try:
                    with open(csv_path, 'a', newline='',encoding="utf-8") as csvfile: #evan- added encoding utf8
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writerow({'Path':newName, 'reportNumber': local_report_number, 'downloadNumber': report_number, 'ExcelPath': excel_folder, 'Date': split.group(1), 'Author': split.group(3),'reportID': split.group(5), 'reportTitle': split.group(4),'cusip6': current_cusip6, 'cusip': current_cusip, 'conm': current_conm, 'tic': current_tic, 'gvkey': current_gvkey, 'timeStamp': time.asctime(), 'is_sp500': current_is_sp500, 'is_sp1500': current_is_sp1500})
                except:
                    with open(csv_path, 'a', newline='') as csvfile:
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        extractError =  'Error extracting data from File Name'
                        writer.writerow({'Path':newName, 'reportNumber': local_report_number, 'downloadNumber': report_number, 'ExcelPath': excel_folder, 'reportTitle': extractError,'cusip6': current_cusip6, 'cusip': current_cusip, 'conm': current_conm, 'tic': current_tic, 'gvkey': current_gvkey, 'timeStamp': time.asctime(), 'is_sp500': current_is_sp500, 'is_sp1500': current_is_sp1500})
def cleanup(Cusip, Year, Tic):
    try:
        global download_directory, desktop_directory, companyErrorNumber
        yearString = str(Year)
        yearString = yearString[2:4]
        cusipTic = str(Cusip) + '_' + Tic
        folderName = os.path.join(desktop_directory, "BetaTest", cusipTic, yearString)
        if os.path.exists(folderName):
            for year in os.listdir(folderName):
                month = os.path.join(folderName, year)
                if(os.path.isdir(month)):
                    for file in os.listdir(month):
                        try:
                            byeBye = os.path.join(folderName, year, file)
                            os.remove(byeBye)
                        except:
                            logging.error("Issue removing file during cleanup")
                elif os.path.isfile(month):
                    try:
                        os.rmdir(month)
                    except:
                        logging.error("Issue removing Excel file during cleanup")
            os.rmdir(folderName)
        for name in os.listdir(download_directory):
            currInfo = str(Year) + '-'
            if( name.startswith(currInfo) ): #Starting with the current search parameters
                adios = os.path.join(download_directory, name)
                os.remove(adios)
            if( name.startswith('Advanced Research Search')):
                adios = os.path.join(download_directory, name)
                os.remove(adios)
        windows = driver.window_handles
        while len(windows) > 1:
            driver.switch_to_window(windows[1])
            driver.close()
            windows = driver.window_handles
        windows = driver.window_handles
        driver.switch_to_window(windows[0])
        logging.info("Cleaned Up for Company Number: " + str(Cusip) + " in Year " + str(Year) + " with Tic: " + str(Tic) )
    except:
        logging.error("ERROR CLEANING UP DURING SEARCH: " + str(Cusip) + " in Year " + str(Year) + " with Tic: " + str(Tic) )
def newDriver():
    fieldnames_skipped = ['cusip6', 'cusip', 'conm', 'tic', 'gvkey',  'timeStamp', 'is_sp500', 'is_sp1500', 'skippedYear', 'timeStamp']
    global driver
    global current_cusip, current_cusip6, current_conm, current_tic, current_gvkey, current_is_sp500, current_is_sp1500, currentCusip, currentTic, currentYear, company
    #currentYear = years[year]
    windows = driver.window_handles
    while True:
        if(len(windows) > 1):
            driver.switch_to_window(windows[0])
            driver.close()
            windows = driver.window_handles
        elif(len(windows) == 1):
            driver.switch_to_window(windows[0])
            driver.close()
            break
        elif(len(windows) == 0):
            break
    driver = webdriver.Chrome(chromedriver, chrome_options=chrome_options)
    logging.info("Getting new ChromeDriver & Restarting Session")
def skip_company():
    fieldnames_skipped = ['cusip6', 'cusip', 'conm', 'tic', 'gvkey',  'timeStamp', 'is_sp500', 'is_sp1500', 'skippedYear', 'timeStamp']
    global current_cusip, current_cusip6, current_conm, current_tic, current_gvkey, current_is_sp500, current_is_sp1500, currentCusip, currentTic, currentYear, company, Year
    logging.info("Skipping Company " + str(currentTic) + ", or: " + str(currentCusip) + " in Year " + str(currentYear))
    with open(skipped_csv_path, 'a', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames_skipped)
        writer.writerow({'cusip6': current_cusip6, 'cusip': current_cusip, 'conm': current_conm, 'tic': current_tic, 'gvkey': current_gvkey, 'is_sp500': current_is_sp500, 'is_sp1500': current_is_sp1500, 'skippedYear': currentYear, 'timeStamp': time.asctime()})
    company = company + 1

#Convert methods for entering date ranges, naming directories, and maintaining consistent formats in the csvFile
def convert(Month, Year): #This is for getting the first of a month and year ; Can Be used in Year Search for December & First Date
    result = "01-"
    if Month == "1-Jan":
            result += "Jan-"
    elif Month == "2-Feb":
            result += "Feb-"
    elif Month == "3-Mar":
            result += "Mar-"
    elif Month == "4-Apr":
            result += "Apr-"
    elif Month == "5-May":
            result += "May-"
    elif Month == "6-Jun":
            result += "Jun-"
    elif Month == "7-Jul":
            result += "Jul-"
    elif Month == "8-Aug":
            result += "Aug-"
    elif Month == "9-Sep":
            result += "Sep-"
    elif Month == "10-Oct":
            result += "Oct-"
    elif Month == "11-Nov":
            result += "Nov-"
    elif Month == "12-Dec":
            result += "Dec-"
    result += Year
    return result
def convert2(Month, Year): #This gets you the last of a month in a year ; Can be used in Year Search for December & Second Date
    result = ""
    if Month == "1-Jan":
            result += "31-" + "Jan-"
    elif Month == "2-Feb":
            result += "28-" + "Feb-"
    elif Month == "3-Mar":
            result += "31-" + "Mar-"
    elif Month == "4-Apr":
            result += "30-" + "Apr-"
    elif Month == "5-May":
            result += "31-" + "May-"
    elif Month == "6-Jun":
            result += "30-" + "Jun-"
    elif Month == "7-Jul":
            result += "31-" + "Jul-"
    elif Month == "8-Aug":
            result += "31-" + "Aug-"
    elif Month == "9-Sep":
            result += "30-" + "Sep-"
    elif Month == "10-Oct":
            result += "31-" + "Oct-"
    elif Month == "11-Nov":
            result += "30-" + "Nov-"
    elif Month == "12-Dec":
            result += "31-" + "Dec-"
    result += Year
    return result
def convert3(Month): #This is to switch "##-MMM" to "##-" for directory creation
    if Month == "1-Jan":
            return '01-'
    elif Month == "2-Feb":
            return '02-'
    elif Month == "3-Mar":
            return '03-'
    elif Month == "4-Apr":
            return '04-'
    elif Month == "5-May":
            return '05-'
    elif Month == "6-Jun":
            return '06-'
    elif Month == "7-Jul":
            return '07-'
    elif Month == "8-Aug":
            return '08-'
    elif Month == "9-Sep":
            return '09-'
    elif Month == "10-Oct":
            return '10-'
    elif Month == "11-Nov":
            return '11-'
    else:
            return '12-'
def convert4(Month, Year): #This gets you the last of a month in a year for Excel naming; Can be used in Year Search for December & Second Date
    result = ""
    if Month == "1-Jan":
            result += "31-" + "01-"
    elif Month == "2-Feb":
            result += "28-" + "02-"
    elif Month == "3-Mar":
            result += "31-" + "03-"
    elif Month == "4-Apr":
            result += "30-" + "04-"
    elif Month == "5-May":
            result += "31-" + "05-"
    elif Month == "6-Jun":
            result += "30-" + "06-"
    elif Month == "7-Jul":
            result += "31-" + "07-"
    elif Month == "8-Aug":
            result += "31-" + "08-"
    elif Month == "9-Sep":
            result += "30-" + "09-"
    elif Month == "10-Oct":
            result += "31-" + "10-"
    elif Month == "11-Nov":
            result += "30-" + "11-"
    elif Month == "12-Dec":
            result += "31-" + "11-"
    result += Year
    return result
def convert5(Month, Year): #This gets you the last of a month in a year for Excel naming; Can be used in Year Search for December & Second Date
    result = ""
    if Month == "1-Jan":
            result += "01-31-"
    elif Month == "2-Feb":
            result += "02-28-"
    elif Month == "3-Mar":
            result += "03-31-"
    elif Month == "4-Apr":
            result += "04-30-"
    elif Month == "5-May":
            result += "05-31-"
    elif Month == "6-Jun":
            result += "06-30-"
    elif Month == "7-Jul":
            result += "07-31-"
    elif Month == "8-Aug":
            result += "08-31-"
    elif Month == "9-Sep":
            result += "09-30-"
    elif Month == "10-Oct":
            result += "10-31-"
    elif Month == "11-Nov":
            result += "11-30-"
    elif Month == "12-Dec":
            result += "12-31-"
    result += Year
    return result
def convert6(Month, Year):
    if Month == 1:
        return '01-01' + Year
    elif Month == 2:
        return '01-02' + Year
    elif Month == 3:
        return '01-03' + Year
    elif Month == 4:
        return '01-04' + Year
    elif Month == 5:
        return '01-05' + Year
    elif Month == 6:
        return '01-06' + Year
    elif Month == 7:
        return '01-07' + Year
    elif Month == 8:
        return '01-08' + Year
    elif Month == 9:
        return '01-09' + Year
    elif Month == 10:
        return '01-10' + Year
    elif Month == 11:
        return '01-11' + Year
    elif Month == 12:
        return '01-12' + Year
    return
def convert7(Month): #This is to switch "##-MMM" to "##-" for directory creation
    if Month == "1-Jan":
            return '01'
    elif Month == "2-Feb":
            return '02'
    elif Month == "3-Mar":
            return '03'
    elif Month == "4-Apr":
            return '04'
    elif Month == "5-May":
            return '05'
    elif Month == "6-Jun":
            return '06'
    elif Month == "7-Jul":
            return '07'
    elif Month == "8-Aug":
            return '08'
    elif Month == "9-Sep":
            return '09'
    elif Month == "10-Oct":
            return '10'
    elif Month == "11-Nov":
            return '11'
    else:
            return '12'
def startYear(Year):
    return "01-Jan-" + str(Year)
def endYear(Year):
    return "31-Dec-" + str(Year)
def quarterStart(quarter, Year):
    if quarter == 0:
        return "01-Jan-" + str(Year)
    if quarter == 1:
        return "01-Apr-" + str(Year)
    if quarter == 2:
        return "01-Jul-" + str(Year)
    if quarter == 3:
        return "01-Oct-" + str(Year)
def quarterEnd(quarter, Year):
    if quarter == 0:
        return "31-Mar-" + str(Year)
    if quarter == 1:
        return "30-Jun-" + str(Year)
    if quarter == 2:
        return "30-Sep-" + str(Year)
    if quarter == 3:
        return "31-Dec-" + str(Year)
def monthStart(Month, Year): #This is for getting the first of a month and year ; Can Be used in Year Search for December & First Date
    result = "01-"
    if Month == 0:
            result += "Jan-"
    elif Month == 1:
            result += "Feb-"
    elif Month == 2:
            result += "Mar-"
    elif Month == 3:
            result += "Apr-"
    elif Month == 4:
            result += "May-"
    elif Month == 5:
            result += "Jun-"
    elif Month == 6:
            result += "Jul-"
    elif Month == 7:
            result += "Aug-"
    elif Month == 8:
            result += "Sep-"
    elif Month == 9:
            result += "Oct-"
    elif Month == 10:
            result += "Nov-"
    elif Month == 11:
            result += "Dec-"
    result += str(Year)
    return result
def monthEnd(Month, Year): #This is for getting the first of a month and year ; Can Be used in Year Search for December & First Date
    result = ""
    if Month == 0:
            result += "31-Jan-"
    elif Month == 1:
            result += "28-Feb-"
    elif Month == 2:
            result += "31-Mar-"
    elif Month == 3:
            result += "30-Apr-"
    elif Month == 4:
            result += "31-May-"
    elif Month == 5:
            result += "30-Jun-"
    elif Month == 6:
            result += "31-Jul-"
    elif Month == 7:
            result += "31-Aug-"
    elif Month == 8:
            result += "30-Sep-"
    elif Month == 9:
            result += "31-Oct-"
    elif Month == 10:
            result += "30-Nov-"
    elif Month == 11:
            result += "31-Dec-"
    result += str(Year)
    return result

if __name__ =='__main__':
    main()
