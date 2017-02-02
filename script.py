#!/usr/bin/python
import sys, getopt, getpass
from datetime import datetime, timedelta
from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection

def main(argv):
    try:
        opts, args = getopt.getopt(argv, "h:u:", ["host=", "user="])
    except getopt.GetoptError:
        print 'script.py -h <host> -u <domain\\username>'
        sys.exit(2)

    host = None
    user = None
    for o, a in opts:
        if o in ("-h", "--host"):
            host = a
        elif o in ("-u", "--user"):
            user = a

    pwd = getpass.getpass("Password for your outlook/exchange account: ")
    if not pwd:
        print 'You need to set the password!'
        sys.exit(2)

    connection = ExchangeNTLMAuthConnection(url=host, username=user, password=pwd)
    service = Exchange2010Service(connection)

    today = datetime.today()
    nextWeek = today + timedelta(days=7)

    events = service.calendar().list_events(start=today, end=nextWeek)

    for event in events.events:
        print "{start} {stop} - {subject}".format(
            start=event.start,
            stop=event.end,
            subject=event.subject
        )

    sys.exit(0)

if __name__ == "__main__":
   main(sys.argv[1:])
