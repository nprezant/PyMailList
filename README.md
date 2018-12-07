# PyMailList

A PyQt5 GUI for sending a single email to a list of recipients.



## Installation

* Clone this repository

```python
git clone https://github.com/nprezant/PyMailList.git
```

* Follow steps 1 and 2 of [Google's instructions](https://developers.google.com/gmail/api/quickstart/python) for how to enable Python requests to the Gmail API. Be sure to save the credentials.json file.
    
* Might as well install all the rest of the package requirements (this should just add PyQt5)

```python
pip install requirements.txt
```



## Background

This project came about because I needed an email list service tool like this for an organization I was running and because I wanted some experience with python GUIs.

The application uses PyQt5 with the googleapiclient and oauth2client packages to manage the emails.



## Example

Example user interface usage.

![DarkThemeExample](/assets/example.gif)



## License

MIT, see [license](/LICENSE.md).



## Acknowledgements

Design themes are a fork of [BreezeStyleSheets](https://github.com/XLTools/BreezeStyleSheets)