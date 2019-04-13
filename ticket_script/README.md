# Ticket Details Extractor

This is a scraper written in python 2.7 for extracting ticket details from a ticketing tool. It can obtain ticket details and 
create a excel report from it and email it to specified users.

## Installation
* Clone the repo
* Create "logs" and "reports" folders inside the root folder of the repo
* Install the dependencies via [pip] (https://pypi.org/project/pip/)
```
pip install -r requirements.txt
```

### Requirements
* OMDb API key ([Register for a key](http://www.omdbapi.com/apikey.aspx))
* Google Chrome web browser (JavaScript enabled)
* Internet connection with a reasonable speed

## Usage
* Enter movie/TV show title you want to look up and click Submit button
* Click the Clear button to clear the textbox fields

## Built With
* [Python 2.7](https://en.wikipedia.org/wiki/HTML5) - Basic markup
* [CSS3](https://en.wikipedia.org/wiki/Cascading_Style_Sheets) - Basic styling
* [Bootstrap 4.1.1](https://getbootstrap.com/) - Responsive framework
* [jQuery 3.3.1](https://jquery.com/) - JS framework

## License
This project is licensed under [MIT](https://choosealicense.com/licenses/mit/) license.

## Acknowledgements
* Uses the [OMDb API](http://www.omdbapi.com/)
* Inspired by Brad Traversy's GitHub finder

## TODO
* Add plot size option
* Format ratings sections
