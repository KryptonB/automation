# Ticket Details Extractor

This is a scraper written in python 2.7 for extracting ticket details from a ticketing tool. It can obtain ticket details and 
create an excel report and email it to specified users.

## Installation
* Clone the repo
* Create **logs** and **reports** folders inside the root folder of the repo
* Install dependencies via [pip](https://pypi.org/project/pip/) package manager
```
pip install -r requirements.txt
```
* Change the python executable path to point to your python installation folder in **_ticket_scraper.bat_** file
* Set your credentials for website and mail server details in the **_config**\configs.json_ file and access them via _extract_configs(configsFile)_ function or put them in
environment variables and access them via _extract_configs_using_env_variables()_ function inside **_ticket_scraper.py_** script

### Requirements
* Python 2.7 (works with python 3.4 also)
* [Selenium web driver](https://sites.google.com/a/chromium.org/chromedriver/) for Google Chrome (already included in **drivers** folder)
* Required 3rd party modules are mentioned in the _requirements.txt_ file
* Internet connection with a reasonable speed

## Usage
* Double click _ticket_scraper.bat_ file and let it run. (It will take around 4 to 5 minutes to complete)
* Optionally you can use [PyInstaller](https://www.pyinstaller.org/) to create a standalone executable (.exe file) and run it

## Built With
* [Python 2.7](https://www.python.org/download/releases/2.7/) - Scripting language
* [Selenium](https://www.seleniumhq.org/) - Browser automation
* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) - For Excel reports
* [Batch Script](https://en.wikipedia.org/wiki/Batch_file) - For scheduling via Windows Task Scheduler

## License
This project is licensed under [MIT](https://choosealicense.com/licenses/mit/) license.

## Acknowledgments
* Many similar apps on the internet

## TODO
* Read the sender list for emails via a file
