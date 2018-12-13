import time
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait


# Sleep times to pause the script (given in seconds)
shortInterval = 3
mediumInterval = 15
longInterval = 40

# Selenium Chrome driver location
chromeDriverPath = 'C:\\Python27\\Custom_tools\\chromedriver.exe'

# URL of the page
gossipLURL = 'https://www.gossiplankanews.com/2018/12/final-decision.html'
# gossipLURL = 'https://www.gossiplankanews.com/2018/12/steave-rixon-for-sri-lanka-cricket.html'
# gossipLURL = 'https://www.gossiplankanews.com/2018/10/arjuna-dematagoda-update.html'

# Comments vote up / vote down button XPaths list
commentXPaths = ['//*[@id="IDComment1067220387"]/div[1]/div/div/a[2]', '//*[@id="IDComment1067220701"]/div[1]/div/div/a[2]', '//*[@id="IDComment1067222914"]/div[1]/div/div/a[2]', '//*[@id="IDComment1067220381"]/div[1]/div/div[2]/a[2]']


# Emoji icon XPaths
kujeethaiIconXPath = '//*[@id="Blog1"]/div[3]/div/div[1]/div[2]/div[6]/div/div[1]/img'
supiriIconXPath = '//*[@id="Blog1"]/div[3]/div/div[1]/div[2]/div[1]/div/div[1]/img'


def commentBot():
    while True:
        browser = webdriver.Chrome(chromeDriverPath)
        browser.delete_all_cookies()
        browser.get(gossipLURL)
        time.sleep(longInterval)
        
        # Remove "idc-disabled" class from the already voted comments
        browser.execute_script("""
            var elements = document.getElementsByClassName('idc-disabled');
            while(elements.length > 0){
                elements[0].classList.remove('idc-disabled');
            }
        """)
        time.sleep(shortInterval)
        
        # Loop thru the selected comments and click their desired button
        for commentXPath in commentXPaths:
            commentIcon = browser.find_element_by_xpath(commentXPath)
            commentIcon.click()
            time.sleep(shortInterval)

        browser.quit()

# Function to click on the selected emoji      
def emojiBot(icon):
    while True:
        browser = webdriver.Chrome(chromeDriverPath)
        browser.get(gossipLURL)
        time.sleep(mediumInterval)
        
        myIcon = browser.find_element_by_xpath(icon)
        myIcon.click()
        time.sleep(shortInterval)
        
        browser.quit()
        

#commentBot()

emojiBot(kujeethaiIconXPath)
