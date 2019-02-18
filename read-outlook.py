#Necessary Installations: pypiwin32, python-dateutil
#SO reference: http://stackoverflow.com/questions/21477599/read-outlook-events-via-python
# https://msdn.microsoft.com/en-us/library/office/ff869026(v=office.15).aspx

# https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff869026(v=office.14)
# https://github.com/bexway/outloook-calendar-reader

import win32com.client, datetime
from dateutil.parser import *
from dateutil.relativedelta import relativedelta
import re
import csv
import calendar
import argparse

def get_args():
    parser = argparse.ArgumentParser(description='Calculate calendar allocation.')
    parser.add_argument('-c', dest='measured_category', default='blocker', help='What CATEGORIES field to calculate')
    parser.add_argument('--exclude', dest='exclude', default='', help='What CATEGORIES field to calculate')
    parser.add_argument('-s', dest='startDate', help='Start date mm/dd/yyyy')
    parser.add_argument('-e', dest='endDate', help='End date mm/dd/yyyy')
    return parser.parse_args()


def main():
    args = get_args();

    print "Accessing Outlook Calendar, please wait..."
    #Access Outlook and get the events from the calendar
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items

    #Sort the events by occurence and then include recurring events
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"

    # get user input range of dates to process
    parsedInput = parse(args.startDate).date()
    begin = parsedInput.strftime("%m/%d/%Y")
    parsedInput = parse(args.endDate).date()
    end = parsedInput.strftime("%m/%d/%Y")

    # restrict appointments to specified range
    appointments = appointments.Restrict("[Start] >= '" +begin+ "' AND [END] <= '" +end+ "'")

    #Generate a dictionary; I need to track appointment dates to count them
    appointmentDictionary = {}
    #Create a regex for time and Subject
    timeregex = re.compile('\d\d/\d\d/\d\d')
    nameregex = re.compile(u'[Nn]ame: ?(?P<name>[\( \)\&;\w]*)', re.UNICODE)
    locationregex = re.compile(u'[Ll]ocation: ?(?P<location>[\( \)\&;\d]*)', re.UNICODE)
    #Note to self: get names from invitees?

    all_events_duration = 0
    target_duration = 0
    for a in appointments:
        #grab the date from the meeting time
        meetingDate = str(a.Start)
        categories = str(a.Categories.encode("utf8"))
        subject = str(a.Subject.encode("utf8"))
        # body = str(a.Body.encode("utf8"))
        duration = str(a.duration)
        
        date = parse(meetingDate).date()
        time = parse(meetingDate).time()

        duration = int(duration)
        if args.exclude not in categories:
            # print duration 
            # print categories
            all_events_duration += duration
            if args.measured_category in categories:
                target_duration += duration
    
    print args.measured_category + " {}:{}".format(target_duration/60, target_duration%60)
    print "All event" + " {}:{}".format(all_events_duration/60, all_events_duration%60)

if __name__ == "__main__":
    main()
