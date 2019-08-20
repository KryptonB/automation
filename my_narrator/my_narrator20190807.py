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




# Wait intervals to be used later in the script until page is loaded in slow networks
tinyPause = 2
longPause = 10
megaPause = 60
shortPause = 5
longerPause = 20


# Set datestamp and timestamp to be used in file names
now = datetime.datetime.now()
datestamp = now.strftime('%Y%m%d_%H%M')
timestamp = now.strftime('%Y%m%d%H%M%S')
datestamp2 = now.strftime('%Y/%m/%d %H:%M')


# Hour of the day to decide which greeting to use
myHour = int(now.hour)


# Web browser object
browser = ''


# Credentials for Artologik
username = ''
password = ''


# URL of Artologik
artoURL = ''


# Initiate speech engine and set the rate of speak
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




#########################################################################
#  Function to initiate browser  and start narrator
#########################################################################
def initiate_narrator(url, username, password):
    global browser
    global lastEmailColumns
    global lastEmailReceivedTime
    global lastEmailReceivedTimeStr
    
    try:
        # Check the driver path and file
        if not (check_file_and_folder_paths(chromeDriverPath)):
            print(chromeDriverPath + ' Does not exist')
            logging.error(chromeDriverPath + ' Does not exist!')
            logging.info('Terminating the program...')
            sys.exit()
            
        else:
            # Initiate a new browser instance and load Artologik
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
        # Find the username textbox and type the username in to it
        u = browser.find_element_by_id('UserName')
        u.click()
        u.send_keys(username)
        logging.info('Username box found and username entered')
        
        time.sleep(tinyPause)
        
        # Find the password textbox and type the password in to it
        p = browser.find_element_by_id('PassWord')
        p.click()
        p.send_keys(password)
        logging.info('Password box found and password entered')
        
        # Find the login button and click it
        login = browser.find_element_by_xpath('/html/body/div/form/p[4]/input')
        login.click()
        logging.info('Login button found')

        # Now we are inside Artologik. Wait 20 secs till the page is fully loaded 
        time.sleep(longerPause)
        
        # Greet the user
        myGreeting = 'Hello, ' + get_greeting(myHour)
        narrate_text(myGreeting)

        # Find the left navigation pane's iframe and switch to that iframe
        leftNavigationFrame = browser.find_element_by_id('Page3')
        browser.switch_to_frame(leftNavigationFrame)
        logging.info('Switched to left navigation menu')
        
        # Find the "Tickets" menu button in the left pane and click it
        ticketsMenuButton = browser.find_element_by_id('td2')
        ticketsMenuButton.click()
        logging.info('Tickets button found and clicked')
        
        # Wait for 20 seconds till the tickets pane is fully loaded
        time.sleep(longerPause)
        
        # Leave the left menu iframe and come back to the default frame
        browser.switch_to_default_content()

        # Find the top pane's iframe and switch to that iframe
        topFrame = browser.find_element_by_id('Set2')
        browser.switch_to_frame(topFrame)
        logging.info('Switched to top pane')
        
        # Find the dropdown list's frame and switch to that iframe
        menuFrame = browser.find_element_by_id('Page2')
        browser.switch_to_frame(menuFrame)
        logging.info('Switched to dropdown list pane')
        
        time.sleep(tinyPause)

        # Find the dropdown list (select box) and click it
        tlist = browser.find_element_by_id('ListID')
        tlist.click()
        logging.info('Dropdown list found and clicked!')
        
        # Find the 'Incoming e-mail' option from the dropdown list and click it
        allNew_and_Current = browser.find_element_by_xpath('//*[@id="ListID"]/option[12]')
        allNew_and_Current.click()
        logging.info('Incoming e-mail option from the list is found and clicked')
        
        # Click the dropdown list again to hide the list items
        tlist.click()

        # Wait 10 secs until incoming mails are loaded.
        time.sleep(longPause)

        # Leave the top pane and come back to the default frame
        browser.switch_to_default_content()

        # Find the iframe where ticket details are displayed and switch to that frame
        incomingTicketFrame = browser.find_element_by_id('Page4')
        browser.switch_to_frame(incomingTicketFrame)
        logging.info('Switched to tickets pane')

        # Wait for 10 secs until all ticket details are loaded
        time.sleep(longPause)
        
        # Find all incoming email rows
        ticketRows = browser.find_elements_by_xpath(ticketRowsXpath)
        logging.info('Reading ticket details...')
        
        # Store the number of emails in the inbox in a variable. 
        # This variable will allow us to make decision in latter steps
        foundCount = len(ticketRows)
    
        if (foundCount != 0):
            # If we come here, then that means there are mails in the inbox            
            # Therefore, tell the user that there are mails in the inbox
            foundEmailPhrase = 'There are, ' + str(foundCount) + ', emails in inbox'
            narrate_text(foundEmailPhrase)
            logging.info(str(foundCount) + ' email(s) found in inbox')
            
            # Since there are emails in the inbox, we will store the received time of last email (that is the latest email)
            lastEmailColumns = ticketRows[-1].find_elements_by_tag_name('td')
            
            # Get the received time using the 2nd <td> tag of the row
            lastEmailReceivedTimeStr = lastEmailColumns[1].text
            
            # Convert that date string to a proper datetime object so that it can be used in comparisons in next steps
            lastEmailReceivedTime = datetime.datetime.strptime(lastEmailReceivedTimeStr, '%Y-%m-%d %H:%M')
            
        else:
            # There are no emails in the inbox. Hence tell that to the user
            narrate_text(emptyInboxPhrase)
            logging.info('There are no emails in inbox')
            
            # Also set the received time of last received mail to a very old time
            # That will allow us to determine if a mail comes during the refresh cycles
            lastEmailReceivedTime = datetime.datetime.strptime(lastEmailReceivedTimeStr, '%Y-%m-%d %H:%M')
            
        
        # Initial Artologik loading is completed. Now we wait 60 seconds to start the refresh loop
        time.sleep(megaPause) 
 
 
        # Loop for refreshing the inbox
        while True:
        
            # Call the refresh function with below arguments
            refresh_inbox(browser, topFrame, menuFrame, incomingEmailIconXpath, incomingTicketFrame, ticketRowsXpath)
            
            # Wait 60 secs before the next refresh
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




#########################################################################   
#  Function for narrating the given text phrase
#########################################################################
def narrate_text(text):
    engine.say(text)
    engine.runAndWait()



    
#########################################################################
#  Function to decide which greeting to use based on hour of day
#########################################################################
def get_greeting(myHour):
    if myHour < 12:
        return 'Good morning'
    elif myHour < 17:
        return 'Good afternoon'
    else:
        return 'Good evening'




#########################################################################
#  Function to extract configuration details from environment variables
#########################################################################
def extract_configs_using_env_variables():
    global artoURL
    global username
    global password
    
    # Get Artologik URL from environment variable
    artoURL = os.environ.get('ARTO_URL')

    # Get the Artologik username and password from environment variables
    username = os.environ.get('ARTO_USERNAME')
    password = os.environ.get('ARTO_PASSWORD')




#########################################################################   
#  Utility function to check file paths and folders
#########################################################################
def check_file_and_folder_paths(filePath):

    # Get the file name from the given path
    givenFileName = os.path.basename(filePath)
    
    # Get the folder name from the given path
    givenFolderName = os.path.dirname(filePath)
    
    if not os.path.exists(givenFolderName):
        # Folder does not exist
        print('Folder: ' + givenFolderName + ' Does not exist!')
        logging.error('Folder: ' + givenFolderName + ' Does not exist!')
        return False
        
    elif not os.path.exists(filePath):
        # File does not exist
        print('File: ' + givenFileName + ' Does not exist!')
        logging.error('File: ' + givenFileName + ' Does not exist!')
        return False
        
    else:
        # Both file and folder exist. Therefore, no issue!
        return True  
  



#########################################################################    
#  Function to refresh the inbox by clicking the mail icon in top pane
#########################################################################
def refresh_inbox(browser, topFrame, menuFrame, incomingEmailIconXpath, incomingTicketFrame, ticketRowsXpath):

    # Leave the incoming tickets iframe and come back to the default frame
    browser.switch_to_default_content()

    # Find the top pane's iframe and switch to that iframe
    browser.switch_to_frame(topFrame)
    logging.info('Switched to top pane. now we are inside refresh_inbox method')
    
    # Switch to top menu iframe
    browser.switch_to_frame(menuFrame)
    logging.info('Switched to dropdown list pane')
    
    time.sleep(tinyPause)

    # Find the incoming mails icon and click it
    incomingEmailIcon = browser.find_element_by_xpath(incomingEmailIconXpath)
    incomingEmailIcon.click()
    print(datetime.datetime.now().strftime('%Y-%d-%m %H:%M:%S') + ' - Incoming mails are loaded')
    logging.info('Incoming mail icon found and clicked')

    time.sleep(tinyPause)
    
    # Leave the top menu iframe and come back to the default frame
    browser.switch_to_default_content()
    
    # Move to the incoming emails pane
    browser.switch_to_frame(incomingTicketFrame)
    logging.info('Switched to tickets pane')

    # Check if there are any mails in the ticket pane and if so, 
    # check if they are new mails or not
    new_mail_checker(ticketRowsXpath)

 


#########################################################################  
#  Function to check if a new mail has arrived
#########################################################################
def new_mail_checker(ticketRowsXpath):
    global lastEmailReceivedTime
    global lastEmailReceivedTimeStr

    # Check if there are any mails in the inbox and if so, put them in a list
    ticketRowsNewer = browser.find_elements_by_xpath(ticketRowsXpath)
    
    # Store the number of emails in the inbox
    foundCountNewer = len(ticketRowsNewer)
    
    if (foundCountNewer != 0): 
        # Since there are emails in the inbox, we will store the received time of last email (that is the latest email)
        # and check if this is a new email or an old one
        lastEmailColumnsNewer = ticketRowsNewer[-1].find_elements_by_tag_name('td')
        lastEmailReceivedTimeNewerStr = lastEmailColumnsNewer[1].text
        lastEmailReceivedNewerTime = datetime.datetime.strptime(lastEmailReceivedTimeNewerStr, '%Y-%m-%d %H:%M')
               
        # Check if the time of the last email in this new run is greater than the older last email time
        if (lastEmailReceivedNewerTime > lastEmailReceivedTime):
            # If we come here, that means this is a newer mail,
            # So, from now on, this becomes our latest mail
            # Therefore, set that time accordingly
            lastEmailReceivedTime = lastEmailReceivedNewerTime
            
            # Extract the sender email address from the newer email
            # 'Support <sas.surveillance@tradetechconsulting.com>' ==> 'sas.surveillance@tradetechconsulting.com'
            senderEmailAddress = lastEmailColumnsNewer[2].text
            senderEmailAddress = senderEmailAddress.split('<')[1].split('>')[0]
            
            # Extract the subject of the newer mail using the 6th <td> tag of the last email row
            emailSubject = lastEmailColumnsNewer[5].text        
            
            # Tell the user that there is a new mail and it's details
            narrate_text(newEmailPhrase)
            narrate_text('Received time, ' + lastEmailReceivedTimeNewerStr + ', Sent by, ' + senderEmailAddress + ', Email subject, ' + emailSubject)
            



# Extract config details via environment variables 
extract_configs_using_env_variables()

# Start the narrator
initiate_narrator(artoURL, username, password)