#!/usr/bin/env python
# See my post on the script here: http://n8henrie.com/

# I got help from here:
# http://bit.ly/Z4PoR3
# and here: http://bit.ly/Z4Plov
# and here: http://bit.ly/Z4PmIU
# and here http://bit.ly/110xGfV

import csv
from icalendar import Calendar, Event, LocalTimezone
from datetime import datetime, timedelta
from random import randint
import easygui
from sys import exit
from os.path import expanduser,isdir

def CheckHeaders(headers):
    '''Makes sure that all the headers are exactly
    correct so that they'll be recognized as the
    necessary keys.'''

    valid_keys = ['End Date', 'Description',
    'All Day Event', 'Start Time', 'Private',
    'End Time', 'Location', 'Start Date', 'Subject']

    for header in headers:
        if header not in valid_keys:
            easygui.msgbox('Looks like one or more of your headers is not quite right, so the script will exit after this. Make sure there aren\'t any leading or trailing spaces, and check the capitalization, and try again. The headers need to be *exactly* like this (without the quotes):\n\n{0}\n\nLooks like the first problematic header was: "{1}".'.format('"' + '" "'.join(valid_keys) + '"', header))

            exit(1)
        else:
            pass

def CleanSpaces(csv_dict):
    '''Cleans trailing spaces from the dictionary
    values, which can break my datetime patterns.'''
    clean_row = {}
    for row in csv_dict:
        for k, v in row.items():
            if k:
                clean_row.update({ k: v.strip()})
        yield clean_row

# The skipinitialspace option eliminates leading
# spaces and reduces blank cells to '' while the CleanSpaces
# function gets rid of trailing spaces.
try:
    if isdir(expanduser("~/Desktop")):
        reader_builder = list(csv.DictReader(open(easygui.fileopenbox(msg="Please select the .csv file to be converted to .ics", title="", default=expanduser("~/Desktop/"), filetypes=["*.csv"]), 'rb'), skipinitialspace = True))
    else:
        reader_builder = list(csv.DictReader(open(easygui.fileopenbox(msg="Please select the .csv file to be converted to .ics", title="", default=expanduser("~/"), filetypes=["*.csv"]), 'rb'), skipinitialspace = True))

# For testing comment 4 lines above (2 x if / else) and use this:
#        reader_builder = list(csv.DictReader(open('path_to_tester.csv', 'rb'), skipinitialspace = True))

except:
    easygui.msgbox('Looks like there was an error opening the file, didn\'t even make it to the conversion part. Sorry!')
    exit(1)

# Filter out events with empty subjects, a required element
# for a calendar event.
# Code found here: http://bit.ly/Z4Pg4h
reader_builder[:] = [d for d in reader_builder if d.get('Subject') != '']

headers = reader_builder[0].keys()
CheckHeaders(headers)

reader = CleanSpaces(reader_builder)

# Start calendar file
cal = Calendar()
cal.add('prodid', 'n8henrie.com')
cal.add('version', '2.0')

# Write the clean list of dictionaries to events.
try:
    rownum = 0
    for row in reader:
        event = Event()
        event.add('summary', row['Subject'])

    # If marked as an "all day event," ignore times.
    # If start and end date are the same
    # or if end date is blank default to a single 24-hour event.
        if row['All Day Event'].lower() == 'true':
            event.add('transp', 'TRANSPARENT')
            event.add('dtstart', datetime.strptime(row['Start Date'], '%m/%d/%Y' ).date())
    #            pdb.set_trace()
            if row['End Date'] == '':
                event.add('dtend', (datetime.strptime(row['Start Date'], '%m/%d/%Y' ) + timedelta(days=1)).date())
            else:
                event.add('dtend', (datetime.strptime(row['End Date'], '%m/%d/%Y' ) + timedelta(days=1)).date())

    # Continue processing events not marked as "all day" events.
        else:

    # Allow either 24 hour time or 12 hour + am/pm
            if row['Start Time'][-2:].lower() in ['am','pm']:
                event.add('dtstart', datetime.strptime(row['Start Date'] + row['Start Time'], '%m/%d/%Y%I:%M %p' ))
            else:
                event.add('dtstart', datetime.strptime(row['Start Date'] + row['Start Time'], '%m/%d/%Y%H:%M' ))

    # Allow blank end dates (assume same day)
            if row['End Date'] == '':
                row['End Date'] = row['Start Date']

            if row['End Time'][-2:].lower() in ['am','pm']:
                event.add('dtend', datetime.strptime(row['End Date'] + row['End Time'], '%m/%d/%Y%I:%M %p' ))
            else:
                event.add('dtend', datetime.strptime(row['End Date'] + row['End Time'], '%m/%d/%Y%H:%M' ))

        event.add('description', row['Description'])
        event.add('location', row['Location'])
        event.add('dtstamp', datetime.replace( datetime.now(), tzinfo=LocalTimezone() ))
        event['uid'] = str(randint(1,10**30)) + datetime.now().strftime('%Y%m%dT%H%M%S') + '___n8henrie.com'

        cal.add_component(event)
    rownum += 1

except:
    if rownum > 0:
        easygui.msgbox('I had a problem with an event. I think I might have gotten through about {0} events and had trouble with an event with subject: {1}. Sorry!'.format(rownum, row['Subject']))
    elif rownum == 0:
        easygui.msgbox('Looks like I didn\'t even get through the first event. Sorry!')
    else:
        easygui.msgbox('Somehow it looks like I processed negative events... that shouldn\'t have happened. Sorry!')
    exit(2)

try:
    # Write final .ics file to same directory as input file.
    if isdir(expanduser('~/Desktop')):
        f = open(easygui.filesavebox(msg='Save .ics File', title='', default=expanduser('~/Desktop/') + 'calendar.ics', filetypes=['*.ics']), 'wb')
    else:
        f = open(easygui.filesavebox(msg='Save .ics File', title='', default=expanduser('~/') + 'calendar.ics', filetypes=['*.ics']), 'wb')

# For testing comment 4 lines above (2 x if / else) and use this:
#    f = open('path_to_tester.csvcalendar.ics', 'wb')

    f.write(cal.to_ical())
    f.close()

except:
    easygui.msgbox('Looks like the conversion went okay, but there was some kind of error writing the file. Sorry!')
    exit(3)
