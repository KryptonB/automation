# Ticket Details Extractor  

[![GitHub](https://img.shields.io/github/license/kryptonb/automation.svg)](https://choosealicense.com/licenses/mit/)

This is a scraper written in python 2.7 for extracting ticket details from a ticketing tool. It can obtain ticket details and 
create an excel report and email it to specified users. There is also a .bat file included in case you want to schedule it to run periodically throughout the day.  
  
Screenshots:  
  
  
![Summary tab](https://github.com/KryptonB/automation/blob/master/ticket_script/screenshots/summary.PNG)  
  
![Ticket details tab](https://github.com/KryptonB/automation/blob/master/ticket_script/screenshots/data.PNG)  
  

## Installation
* Clone the repo
* Create two subfolders called **logs** and **reports** inside the root folder of the repo. **logs** folder will hold the log files and **reports** folder will hold
the generated reports
* Install dependencies via [pip](https://pypi.org/project/pip/) package manager
```
pip install -r requirements.txt
```
* Change the python executable path to point to your python installation folder in [ticket_script.bat](ticket_script.bat) file
* Set your credentials for website and mail server details in the [configs.json](config/configs.json) file and access them via _extract_configs(configsFile)_ function or put them in
environment variables and access them via _extract_configs_using_env_variables()_ function inside [ticket_script.py](ticket_script.py) script

### Requirements
* Python 2.7 (works with python 3.4 also)
* [Selenium web driver](https://sites.google.com/a/chromium.org/chromedriver/) for Google Chrome (already included in **drivers** folder)
* Required 3rd party modules are mentioned in the [requirements.txt](requirements.txt) file
* Internet connection with a reasonable speed

## Usage
* Double click [ticket_script.bat](ticket_script.bat) file and let it run. (It will take around 4 to 5 minutes to complete)
* Optionally you can use [PyInstaller](https://www.pyinstaller.org/) to create a standalone executable (.exe file) and run it

## Built With
* [Python 2.7](https://www.python.org/download/releases/2.7/) - Scripting language
* [Selenium](https://www.seleniumhq.org/) - Browser automation
* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) - For Excel reports
* [Batch Script](https://en.wikipedia.org/wiki/Batch_file) - For scheduling via Windows Task Scheduler

## Acknowledgments
* Many similar apps on the Internet

## TODO
* Read the sender list for emails via a file
