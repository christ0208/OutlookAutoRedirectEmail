# Outlook Auto-configure Redirect Email

The project purpose is for quick and easy setting for redirection of source's Microsoft email to target mail. 
Just one email can be handled manually, but if 10 or more email need to be configured, it would take some time and exhaust yourself.

## Getting Started

The project has been programmed using Python language, Python 3.6 to be exact. 

### Prerequisites
- Install Python 3.6 (Other version of Python is fine, as long as you can install Selenium).
- Install selenium package.
```python
    pip install selenium
```
- Install xlrd package for reading excel file as the data source.
```python
    pip install xlrd
```
- Prepare Chrome Driver for support selenium in giving view of automated process.
You can download it [here](https://chromedriver.chromium.org/downloads). After download, please put chromedriver.exe in one level with this project.

## Deployment

- Duplicate data-example.xlsx and rename it to data.xlsx
- Fill data.xlsx with several Microsoft Emails you want to configure the email redirection.
- Run Python script using following command

```python
    python main.py
```
- Voilà. Those email you list is being configured.

## Built With

* [Python 3.6](https://www.python.org/downloads/release/python-368/) - Main language used.

## Contributing

Please read [CONTRIBUTING.md]() for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](). 

## Authors

* **Christopher L** - *Initial work* - [christ0208](https://github.com/christ0208)

See also the list of [contributors]() who participated in this project.

## Acknowledgments

* Adobe Auto-configure using Selenium by HR.
* https://selenium-python.readthedocs.io/
* 
