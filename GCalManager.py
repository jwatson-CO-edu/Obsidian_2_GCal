import os, datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

_SECRET_GCAL_KEY = "keys/google_calendar.json"
_SESSION_TOKEN   = "token.json"
# If modifying these scopes, delete the file token.json.
_SCOPES = ["https://www.googleapis.com/auth/calendar.events",] # View and edit events on all your calendars.

class GCalManager:
    """ Handles Google Calendar Operations """
    def __init__( self, secretFile : str = _SECRET_GCAL_KEY, sessionToken : str = _SESSION_TOKEN ):
        self.creds   = None
        self.flow    = None
        self.service = None
        if os.path.exists( sessionToken ):
            self.creds = Credentials.from_authorized_user_file( sessionToken, _SCOPES )
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                self.flow = InstalledAppFlow.from_client_secrets_file(
                    secretFile, _SCOPES
                )
            self.creds = self.flow.run_local_server( port = 0 )
            # Save the credentials for the next run
            with open( sessionToken, "w" ) as token:
                token.write( self.creds.to_json() )
        try:
            self.service = build( "calendar", "v3", credentials = self.creds )
        except HttpError as error:
            raise RuntimeError( f"An error occurred!: {error}" )
        

    def get_N_last_events( self, N = 10 ) -> list[dict]:
        """ Call the Calendar API and get the last `N` events """
        now = datetime.datetime.now(tz=datetime.timezone.utc).isoformat()
        print("Getting the upcoming 10 events")
        events_result = (
            self.service.events().list(
                calendarId   = "primary",
                timeMin      = now,
                maxResults   = N,
                singleEvents = True,
                orderBy      = "startTime",
            ).execute()
        )
        return events_result.get( "items", [] )
    

    def create_event( self,
                      title : str = "Action Item -or- Meeting",
                      loc   : str = "Your Desk",
                      desc  : str = "Action Item -or- Meeting, Please do this!",
                      ):
        """ Add an event to GCal """
        event = {
            'summary': title,
            'location': loc,
            'description': desc,
            'start': {
                'dateTime': '2015-05-28T09:00:00-07:00',
                'timeZone': 'America/Los_Angeles',
            },
            'end': {
                'dateTime': '2015-05-28T17:00:00-07:00',
                'timeZone': 'America/Los_Angeles',
            },
            'recurrence': [
                'RRULE:FREQ=DAILY;COUNT=2'
            ],
            'attendees': [
                {'email': 'lpage@example.com'},
                {'email': 'sbrin@example.com'},
            ],
            'reminders': {
                'useDefault': False,
                'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
                ],
            },
        }

        event = self.service.events().insert( calendarId = 'primary', body = event).execute()
        print( f"Event created: {event.get('htmlLink')}" )
