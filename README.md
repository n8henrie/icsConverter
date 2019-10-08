icsConverter
============

## About icsConverter
It can be a major hassle to take a spreadsheet of events and import them to a calendar. This Python script intends to make that process easier by converting a .csv file to an .ics file that can be imported by most modern calendaring apps.

The spreadsheet must be meticulously formatted and exported to the .csv. If I'm going to be honest, the formatting is the hard part; the task handled by this app can be done easily enough by uploading the .csv to Google Calendar and re-downloading as .ics. However, for those that would rather not use Google Calendar for the conversion, this should do the trick.

### ChangeLog
Hosted separately, [here](https://github.com/n8henrie/icsConverter/blob/master/ChangeLog.md).

## Installation
#### OSX
I've used [py2app](https://pypi.python.org/pypi/py2app/) to create a regular old OSX app that *should* work on 10.6+

#### As a Python script (Linux?)
If you're choosing to run it as a script, you probably know a lot more about Python than I do. In case this isn't true... you can see the list of required modules in requirements.txt. If you have pip installed, you can install these in one go using `pip install -r requirements.txt`

#### Windows
Hmmm. Unclear what it would take to port to windows. You may give a shot to the webapp.

#### WebApp

icsConverterWebapp used to be hosted at http://icsconverterwebapp.n8henrie.com,
but not anymore. The updated version is [live at
icw.n8henrie.com](http://icw.n8henrie.com), source code at
<https://github.com/n8henrie/icw>.

## Usage
Not going to write this over again. See my post [here](http://n8henrie.com/2013/05/spreadsheet-to-calendar/).

## GitHub Pages
Giving a shot to a [GitHub Page for icsConverter](http://n8henrie.github.io/icsConverter/).
