from __future__ import print_function
import datetime
import os.path
import sys
import tabula
from datetime import datetime
from dateutil.relativedelta import *
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

class TimetableGenerator():
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/calendar']
        self.service = self.login()
        self.course = sys.argv[1]
        self.calendarID = None
        self.weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        self.semesterStartDate = sys.argv[2]

    def login(self):
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        return build('calendar', 'v3', credentials=creds)

    def getCalendars(self):
        """Returns the list of existing calendars\n
            calendar = {
                kind: str,
                etag: str,
                id: str,
                summary: str,
                timeZone: str,
                colorId: int,
                backgroundColor: str,
                foregroundColor: str,
                selected: bool,
                accessRole: str,
                defaultReminders: str[],
                conferenceProperties: {
                    allowedConferenceSolutionTypes: str[]
                }
            }
        """
        return self.service.calendarList().list(pageToken=None).execute()['items']

    def createCalendar(self, name: str, timeZone: str="Europe/Berlin"):
        """Creates calendar with the name 'name', returns calendar object\n
            calendar = {
                kind: str,
                etag: str,
                id: str,
                summary: str,
                timeZone: str,
                conferenceProperties: {
                    allowedConferenceSolutionTypes: str[]
                }
            }
        """
        calendar = {
            'summary': name,
            'timeZone': timeZone
        }
        return self.service.calendars().insert(body=calendar).execute()

    def createEvent(self, name:str, location:str, timeStart:str, timeEnd:str, recurranceType:str, description:str=None):
        # 2021-10-26T09:00:00
        event = {
            'summary': name,
            'location': location,
            'description': description,
            'start': {
                'dateTime': timeStart,
                'timeZone': 'Europe/Berlin',
            },
            'end': {
                'dateTime': timeEnd,
                'timeZone': 'Europe/Berlin',
            },
            'recurrence': [
                f"RRULE:FREQ={recurranceType};COUNT=100"
            ],
            'reminders': {
                'useDefault': False,
                'overrides': [
                    { 'method': 'popup', 'minutes': 10 },
                ],
            }
        }

        event = self.service.events().insert(calendarId=self.calendarID, body=event).execute()
        print(event)

    def generate(self):

        calendars = self.getCalendars()
    
        try:
            with open('calendars.txt', 'r') as f:
                lines = f.readlines()
                for line in lines:
                    course, id = line.strip().split(',')
                    if course == self.course:
                        self.calendarID = id
        except:
            calendar = self.createCalendar(name=self.course)
            self.calendarID = calendar['id']
            with open('calendars.txt', 'w') as f:
                f.write(f"{self.course},{calendar['id']}")
            
        dfs = tabula.read_pdf(f"{self.course}.pdf", pages='all')

        for df in dfs:
            for slot in df.values:
                course = slot[0]
                semester = slot[1]
                moduleID = slot[2]
                courseName = slot[3].replace('\r', ' ')
                courseType = slot[4]
                lecturer = slot[5]
                courseFormat = slot[6]
                date = slot[7]
                room = slot[8]
                info = slot[9]

                if isinstance(date, float): continue
                date = date.split('\r')

                for time in date:
                    
                    recurrenceType = None
                    kickOff = None
                    kickOffDate = None
                
                    items = time.split(' ')
                    if '-' in items: items.remove('-')
                    for i in range(len(items)):
                        items[i] = items[i].replace('-', "").strip()

                    weekday, timeStart, timeEnd = items[:3]

                    timeStart = timeStart.replace('.', ':')
                    timeEnd = timeEnd.replace('.', ':')

                    extra = items[3:]

                    if weekday in self.weekdays:

                        if 'biweekly' in extra:
                            recurrenceType = 'biweekly'
                        else:
                            recurrenceType = 'weekly'
                        
                            
                        if 'kick' in extra and 'off' in extra:
                            kickOff = True
                            for item in extra:
                                if '.' in item:
                                    if len(item.split('.')[-1]) == 2:
                                        _day, _month, _year = item.split('.')
                                        kickOffDate = "-".join([_day, _month, "20"+_year])
                                    else:
                                        kickOffDate = item.replace('.', '-')
                                    break
                    else:
                        print("Unable to understand:", ", ".join(items))
                        continue
                    
                try:
                    timeStartDate = datetime.strptime(f"{kickOffDate if kickOff else self.semesterStartDate} {timeStart}", '%d-%m-%Y %H:%M')
                    timeEndDate = datetime.strptime(f"{kickOffDate if kickOff else self.semesterStartDate} {timeEnd}", '%d-%m-%Y %H:%M')

                    if not kickOff:
                        timeStartDate = timeStartDate + relativedelta(days=self.weekdays.index(weekday))
                        timeEndDate = timeEndDate + relativedelta(days=self.weekdays.index(weekday))
                    
                    self.createEvent(name=courseName, location=room, timeStart=timeStartDate.isoformat(), timeEnd=timeEndDate.isoformat(), description=f"{courseType} {courseFormat}\n{info}\nSemester: {semester}\nLecturer: {lecturer}\nModule ID: {moduleID}", recurranceType=recurrenceType.upper())
                except Exception as e:
                    print(e)

generator = TimetableGenerator()

generator.generate()