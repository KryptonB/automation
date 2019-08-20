#########################################################################
#
# This script will log into Artologik and keep watching  
# and inform us / narrate us when a new mail comes in.
# It will also readout the details of the mail too.
# 
# It uses python's pyttsx3 module to convert text-to-speech
#    
# Developed by FSS Team
#
#########################################################################

import os
import sys
import time
import json
import pyttsx3
import logging
import datetime

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import UnexpectedAlertPresentException

tinyPause = 2
longPause = 10
megaPause = 60
shortPause = 5
longerPause = 20

now = datetime.datetime.now()
datestamp = now.strftime('%Y%m%d_%H%M')
timestamp = now.strftime('%Y%m%d%H%M%S')
datestamp2 = now.strftime('%Y/%m/%d %H:%M')

myHour = int(now.hour)

browser = ''

username = ''
password = ''

artoURL = ''

engine = pyttsx3.init()
engine.setProperty('rate',140)


# Our WSS clients
clientList = ['volvo', 'sas', 'husqvarna', 'electrolux', 'bmw', 'skanska', 'fmo', 'ericsson', 'sbab']


# List to hold the <td> columns of the last email received
lastEmailColumns = []


# Last email received time
lastEmailReceivedTime = ''
lastEmailReceivedTimeStr = '1900-01-01 01:00'


# Chrome driver path
chromeDriverPath = '.\\drivers\\chromedriver.exe'


# Default log level. Only messages in and above this level will be printed to the log
# Accepted levels are DEBUG, INFO, WARNING, ERROR, CRITICAL
logLevel = logging.INFO


# Log file path and name. Make sure that the 'logs' folder exists. 
# If the 'logs' folder does not exist, you need to create it before running the script
logFile = '.\\logs\\narrator_' + datestamp + '.log'


# Log configurations
logFormat = '[%(asctime)s] %(levelname)s (%(filename)s: %(lineno)d) %(message)s'
logging.basicConfig(filename=logFile, level=logLevel, format=logFormat)


# Phrase for informing the user when there are no emails in inbox
emptyInboxPhrase = 'Hey, Good news!. There are no emails in inbox'


# Phrase for informing the user when there is a new email in inbox
newEmailPhrase = 'New, email received'


# Phrase for informing the user when there is an error in the script
errorPhrase = 'Oops! There was an error, in script. I am, terminating.'


# XPath string for incoming mail icon found at top iframe
incomingEmailIconXpath = '/html/body/form/table/tbody/tr[2]/td[2]/a[2]/span'


# XPath string for ticket rows in the main ticket iframe
ticketRowsXpath = '//*[@id="ctl00_ContentPlaceHolder1_up_rgMailList"]/table/tbody/tr'

def initiate_narrator(url, username, password):
    global browser
    global lastEmailColumns
    global lastEmailReceivedTime
    global lastEmailReceivedTimeStr
    
    try:
        if not (check_file_and_folder_paths(chromeDriverPath)):
            print(chromeDriverPath + ' Does not exist')
            logging.error(chromeDriverPath + ' Does not exist!')
            logging.info('Terminating the program...')
            sys.exit()           
        else:
            browser = webdriver.Chrome(chromeDriverPath)
            browser.get(url)
            print('Loading Artologik...')
            logging.info('Chrome driver path: ' + chromeDriverPath)
            logging.info('Website URL: ' + url)      
    except WebDriverException as e:
        print('Selenium exception!\n' + str(e))
        logging.error(e, exc_info=True)
        logging.info('Terminating the program...')
        narrate_text(errorPhrase)
        sys.exit()
    time.sleep(shortPause)
    
    try:
        u = browser.find_element_by_id('UserName')
        u.click()
        u.send_keys(username)
        logging.info('Username box found and username entered')
        
        time.sleep(tinyPause)
        
        p = browser.find_element_by_id('PassWord')
        p.click()
        p.send_keys(password)
        logging.info('Password box found and password entered')

        login = browser.find_element_by_xpath('/html/body/div/form/p[4]/input')
        login.click()
        logging.info('Login button found')

        time.sleep(longerPause)

        myGreeting = 'Hello, ' + get_greeting(myHour)
        narrate_text(myGreeting)

        leftNavigationFrame = browser.find_element_by_id('Page3')
        browser.switch_to_frame(leftNavigationFrame)
        logging.info('Switched to left navigation menu')

        ticketsMenuButton = browser.find_element_by_id('td2')
        ticketsMenuButton.click()
        logging.info('Tickets button found and clicked')

        time.sleep(longerPause)

        browser.switch_to_default_content()

        topFrame = browser.find_element_by_id('Set2')
        browser.switch_to_frame(topFrame)
        logging.info('Switched to top pane')

        menuFrame = browser.find_element_by_id('Page2')
        browser.switch_to_frame(menuFrame)
        logging.info('Switched to dropdown list pane')
        
        time.sleep(tinyPause)

        tlist = browser.find_element_by_id('ListID')
        tlist.click()
        logging.info('Dropdown list found and clicked!')

        allNew_and_Current = browser.find_element_by_xpath('//*[@id="ListID"]/option[12]')
        allNew_and_Current.click()
        logging.info('Incoming e-mail option from the list is found and clicked')

        tlist.click()
        time.sleep(longPause)
        browser.switch_to_default_content()

        incomingTicketFrame = browser.find_element_by_id('Page4')
        browser.switch_to_frame(incomingTicketFrame)
        logging.info('Switched to tickets pane')

        time.sleep(longPause)

        ticketRows = browser.find_elements_by_xpath(ticketRowsXpath)
        logging.info('Reading ticket details...')
        
        foundCount = len(ticketRows)
    
        if (foundCount != 0):
            foundEmailPhrase = 'There are, ' + str(foundCount) + ', emails in inbox'
            narrate_text(foundEmailPhrase)
            logging.info(str(foundCount) + ' email(s) found in inbox')

            lastEmailColumns = ticketRows[-1].find_elements_by_tag_name('td')

            lastEmailReceivedTimeStr = lastEmailColumns[1].text

            lastEmailReceivedTime = datetime.datetime.strptime(lastEmailReceivedTimeStr, '%Y-%m-%d %H:%M')           
        else:
            narrate_text(emptyInboxPhrase)
            logging.info('There are no emails in inbox')
            lastEmailReceivedTime = datetime.datetime.strptime(lastEmailReceivedTimeStr, '%Y-%m-%d %H:%M')

        time.sleep(megaPause) 

        while True:
            refresh_inbox(browser, topFrame, menuFrame, incomingEmailIconXpath, incomingTicketFrame, ticketRowsXpath)
            time.sleep(megaPause)
    except NoSuchElementException as e:
        print('Element not found!\n' + str(e))
        logging.error(e, exc_info=True)
        logging.info('Terminating the program...')
        browser.quit()
        narrate_text(errorPhrase)
        sys.exit() 
    except UnexpectedAlertPresentException as e:
        print('Username or Password is incorrect!\n' + str(e))
        logging.error(e, exc_info=True)
        logging.info('Terminating the program...')
        browser.quit()
        narrate_text(errorPhrase)
        sys.exit()


def narrate_text(text):
    engine.say(text)
    engine.runAndWait()


def get_greeting(myHour):
    if myHour < 12:
        return 'Good morning'
    elif myHour < 17:
        return 'Good afternoon'
    else:
        return 'Good evening'


def extract_configs_using_env_variables():
    global artoURL
    global username
    global password

    artoURL = os.environ.get('ARTO_URL')
    username = os.environ.get('ARTO_USERNAME')
    password = os.environ.get('ARTO_PASSWORD')
    

def check_file_and_folder_paths(filePath):
    givenFileName = os.path.basename(filePath)
    givenFolderName = os.path.dirname(filePath)
    
    if not os.path.exists(givenFolderName):
        print('Folder: ' + givenFolderName + ' Does not exist!')
        logging.error('Folder: ' + givenFolderName + ' Does not exist!')
        return False    
    elif not os.path.exists(filePath):
        print('File: ' + givenFileName + ' Does not exist!')
        logging.error('File: ' + givenFileName + ' Does not exist!')
        return False
    else:
        return True  
  

def refresh_inbox(browser, topFrame, menuFrame, incomingEmailIconXpath, incomingTicketFrame, ticketRowsXpath):
    browser.switch_to_default_content()
    browser.switch_to_frame(topFrame)
    logging.info('Switched to top pane. now we are inside refresh_inbox method')

    browser.switch_to_frame(menuFrame)
    logging.info('Switched to dropdown list pane')
    
    time.sleep(tinyPause)

    incomingEmailIcon = browser.find_element_by_xpath(incomingEmailIconXpath)
    incomingEmailIcon.click()
    print(datetime.datetime.now().strftime('%Y-%d-%m %H:%M:%S') + ' - Incoming mails are loaded')
    logging.info('Incoming mail icon found and clicked')

    time.sleep(tinyPause)

    browser.switch_to_default_content()

    browser.switch_to_frame(incomingTicketFrame)
    logging.info('Switched to tickets pane')

    new_mail_checker(ticketRowsXpath)

def new_mail_checker(ticketRowsXpath):
    global lastEmailReceivedTime
    global lastEmailReceivedTimeStr

    ticketRowsNewer = browser.find_elements_by_xpath(ticketRowsXpath)
    foundCountNewer = len(ticketRowsNewer)
    
    if (foundCountNewer != 0): 
        lastEmailColumnsNewer = ticketRowsNewer[-1].find_elements_by_tag_name('td')
        lastEmailReceivedTimeNewerStr = lastEmailColumnsNewer[1].text
        lastEmailReceivedNewerTime = datetime.datetime.strptime(lastEmailReceivedTimeNewerStr, '%Y-%m-%d %H:%M')

        if (lastEmailReceivedNewerTime > lastEmailReceivedTime):
            lastEmailReceivedTime = lastEmailReceivedNewerTime
            senderEmailAddress = lastEmailColumnsNewer[2].text
            senderEmailAddress = senderEmailAddress.split('<')[1].split('>')[0]

            emailSubject = lastEmailColumnsNewer[5].text        

            narrate_text(newEmailPhrase)
            narrate_text('Received time, ' + lastEmailReceivedTimeNewerStr + ', Sent by, ' + senderEmailAddress + ', Email subject, ' + emailSubject)

extract_configs_using_env_variables()

initiate_narrator(artoURL, username, password)