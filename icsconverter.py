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
import sys
from os.path import expanduser,isdir
import logging

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

class HeadersError(Exception):
    pass

class DateTimeError(Exception):
    pass

def check_headers(headers):
    '''Makes sure that all the headers are exactly
    correct so that they'll be recognized as the
    necessary keys.'''

    valid_keys = ['End Date', 'Description',
    'All Day Event', 'Start Time', 'Private',
    'End Time', 'Location', 'Start Date', 'Subject']

    if (set(headers) != set(valid_keys)
    or len(headers) != len(valid_keys)):
        for header in headers:
            if header not in valid_keys:
                if header == '':
                    header = '"" (<- an empty column)'
                logger.error('Invalid header: {}'.format(header))
                easygui.msgbox('Looks like one or more of your headers is not quite right, so the script will exit after this. Make sure there aren\'t any leading or trailing spaces, and check the capitalization, and try again. The headers need to be *exactly* like this (without the quotes):\n\n{0}\n\nLooks like the first problematic header was: "{1}".'.format('"' + '" "'.join(valid_keys) + '"', header))
                raise HeadersError('Something isn\'t right with the headers.')
            elif len(headers) < len(valid_keys):
                logger.error('Missing headers: {}'.format(list(set(valid_keys) - set(headers))))
                easygui.msgbox('I think you may be missing the following header(s): {}'.format(list(set(valid_keys) - set(headers))))
                raise HeadersError('Something isn\'t right with the headers.')
            elif len(headers) > len(valid_keys):
                duplicate = list(set([i for i in headers if headers.count(i) > 1]))
                logger.error('Extra headers: {}'.format(duplicate))
                easygui.msgbox('You might have one or more duplicate headers: {}'.format(duplicate))
                raise HeadersError('Something isn\'t right with the headers.')
    else:
        return 'headers passed'

def clean_spaces(csv_dict):
    '''Cleans trailing spaces from the dictionary
    values, which can break my datetime patterns.'''
    clean_row = {}
    for row in csv_dict:
        for k, v in row.items():
            if v:
                clean_row.update({ k: v.strip() })
            else:
                clean_row.update({ k: None })

        yield clean_row

def check_dates_and_times(
        start_date = None, start_time = None,
        end_date = None, end_time = None, all_day = None, subject = None
    ):
    '''Checks the dates and times to make sure everything is kosher.'''

    logger.debug('Date checker started.')

    # Gots to have a start date, no matter what.
    if start_date in ['', None]:
        logger.error('Missing a start date')
        easygui.msgbox('''You're missing a start date. All events need start dates! '''
        '''The Subject for the event was: {}'''.format(subject))
        raise DateTimeError('Missing a start date')
        return False

    for date in [start_date, end_date]:
        if date not in ['', None]:
            try:
                datetime.strptime(date, '%m/%d/%Y')
            except:
                easygui.msgbox('''Problematic date found: {}\n'''
                '''Make sure all dates are MM/DD/YYYY and try again.'''.format(date))
                logger.error('Problem with date formatting. Date: {}'.format(date))
                raise DateTimeError('Something isn\'t right with the dates.')
                return False

    for time in [start_time, end_time]:
        if time not in ['', None]:
            try:
                time = time.replace(' ', '')
                if time[-2:].lower() in ['am','pm']:
                    datetime.strptime(time, '%I:%M%p')
                else:
                    datetime.strptime(time, '%H:%M' )
            except:
                easygui.msgbox('''Problematic time found: {}\n'''
                '''Make sure all times are HH:MM (either 24h or with am / pm) and try again.'''.format(date))
                logger.error('Problem with time formatting. Time: {}'.format(time))
                raise DateTimeError('Something isn\'t right with the times.')
                return False

    if all_day == None or all_day.lower() != 'true':
       if not (start_time and end_time):
           easygui.msgbox('''Missing a required time field in a non-all_day event on date: {}.\n'''
           '''Remember, if it's not an all_day event, you must have both start and end times!'''.format(date))
           logger.error('Missing a required time field in a non-all_day event on date: {}.'.format(start_date))
           raise DateTimeError('Missing a required time field in a non-all_day event.')
           return False

    logger.debug('Date checker ended.')
    return True

def main(infile=None):
    # The skipinitialspace option eliminates leading
    # spaces and reduces blank cells to '' while the clean_spaces
    # function gets rid of trailing spaces.
    try:
        if infile == None:
            start_dir = '~/'
            if isdir(expanduser("~/Desktop")):
                start_dir = '~/Desktop/'
            msg = 'Please select the .csv file to be converted to .ics'
            infile = easygui.fileopenbox(msg=msg, title="", default=expanduser(start_dir), filetypes=["*.csv"])

        reader_builder = list(csv.DictReader(open(infile, 'U'), skipinitialspace = True))

    # For testing comment 4 lines above (2 x if / else) and use this:
    #        reader_builder = list(csv.DictReader(open('path_to_tester.csv', 'rb'), skipinitialspace = True))

    except Exception as e:
        logger.exception(e)
        easygui.msgbox("Looks like there was an error opening the file, didn't even make it to the conversion part. Sorry!")
        sys.exit(1)

    # Filter out events with empty subjects, a required element
    # for a calendar event.
    # Code found here: http://bit.ly/Z4Pg4h
    reader_builder[:] = [d for d in reader_builder if d.get('Subject') != '']

    headers = reader_builder[0].keys()
    logger.debug('reader_builder[0].keys(): {}'.format(headers))
    check_headers(headers)

    reader = clean_spaces(reader_builder)

    # Start calendar file
    cal = Calendar()
    cal.add('prodid', 'n8henrie.com')
    cal.add('version', '2.0')

    # Write the clean list of dictionaries to events.
    rownum = 0
    try:
        for row in reader:
            event = Event()
            event.add('summary', row['Subject'])

            try:
                check_dates_and_times(
                    start_date = row.get('Start Date'),
                    start_time = row.get('Start Time'),
                    end_date = row.get('End Date'),
                    end_time = row.get('End Time'),
                    all_day = row.get('All Day Event'),
                    subject = row.get('Subject')
                )
            except DateTimeError as e:
                sys.exit(e)

        # If marked as an "all day event," ignore times.
        # If start and end date are the same
        # or if end date is blank default to a single 24-hour event.
            if row.get('All Day Event') != None and row['All Day Event'].lower() == 'true':

                # All-day events will not be marked as 'busy'
                event.add('transp', 'TRANSPARENT')

                event.add('dtstart', datetime.strptime(row['Start Date'], '%m/%d/%Y' ).date())

                if row.get('End Date') in ['', None]:
                    event.add('dtend', (datetime.strptime(row['Start Date'], '%m/%d/%Y' ) + timedelta(days=1)).date())
                else:
                    event.add('dtend', (datetime.strptime(row['End Date'], '%m/%d/%Y' ) + timedelta(days=1)).date())

            # Continue processing events not marked as "all day" events.
            else:

                # Events with times should be 'busy' by default
                event.add('transp', 'OPAQUE')

                # Get rid of spaces
                # Note: Must have both start and end times if not all_day, already checked
                row['Start Time'] = row['Start Time'].replace(' ', '')
                row['End Time'] = row['End Time'].replace(' ', '')

                # Allow either 24 hour time or 12 hour + am/pm
                if row['Start Time'][-2:].lower() in ['am','pm']:
                    event.add('dtstart', datetime.strptime(row['Start Date'] + row['Start Time'], '%m/%d/%Y%I:%M%p' ))
                else:
                    event.add('dtstart', datetime.strptime(row['Start Date'] + row['Start Time'], '%m/%d/%Y%H:%M' ))

                # Allow blank end dates (assume same day)
                if row.get('End Date') in ['', None]:
                    row['End Date'] = row['Start Date']

                if row['End Time'][-2:].lower() in ['am','pm']:
                    event.add('dtend', datetime.strptime(row['End Date'] + row['End Time'], '%m/%d/%Y%I:%M%p' ))
                else:
                    event.add('dtend', datetime.strptime(row['End Date'] + row['End Time'], '%m/%d/%Y%H:%M' ))

            if row.get('Description'):
                event.add('description', row['Description'])
            if row.get('Location'):
                event.add('location', row['Location'])

            event.add('dtstamp', datetime.replace( datetime.now(), tzinfo=LocalTimezone() ))
            event['uid'] = str(randint(1,10**30)) + datetime.now().strftime('%Y%m%dT%H%M%S') + '___n8henrie.com'

            cal.add_component(event)
            rownum += 1

    except Exception, e:
        if rownum > 0:
            easygui.msgbox('I had a problem with an event. I think I might have gotten through about {0} events and had trouble with an event with subject: {1}. Sorry!'.format(rownum, row['Subject']))
            logger.exception(e)
        elif rownum == 0:
            easygui.msgbox('Looks like I didn\'t even get through the first event. Sorry!')
            logger.exception(e)
        else:
            easygui.msgbox('Somehow it looks like I processed negative events... that shouldn\'t have happened. Sorry!')
            logger.exception(e)
        sys.exit(2)

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

    except Exception, e:
        easygui.msgbox('Looks like the conversion went okay, but there was some kind of error writing the file. Sorry!')
        logger.exception(e)
        sys.exit(3)

if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            main(infile=sys.argv[1])
        else:
            main()
    except (HeadersError, DateTimeError) as e:
        sys.exit(e)
