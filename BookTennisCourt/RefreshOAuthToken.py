###########################
##### Version 2.0 #########
###########################


import sys
import traceback
from datetime import datetime  
from datetime import timedelta  
from CreateCalendarEvent.cal_setup import get_calendar_service


def create_calendar_event(event_date, event_hour, event_summary):
    try:
       # creates one hour event tomorrow 10 AM IST
       
       print(event_date)
       service = get_calendar_service()
       event_datetime = datetime(event_date.year, event_date.month, event_date.day, event_hour)+timedelta(days=0)
       start = event_datetime.isoformat()
       end = (event_datetime + timedelta(hours=1)).isoformat()
    
       event_result = service.events().insert(calendarId='primary'
          ,body={
                   'summary': event_summary,
                   'description': 'Tennis Court',
                   'start': {"dateTime": start, 'timeZone': 'America/Chicago'},
                   'end': {"dateTime": end, 'timeZone': 'America/Chicago'},         
                   'reminders': {'useDefault': False,
                                 'overrides': [{'method': 'email', 'minutes': 24 * 60},
                                               {'method': 'email', 'minutes': 8 * 60},
                                               {'method': 'popup', 'minutes': 120},
                                              ],
                                }
               }
          ,sendUpdates='all'
       )
       event_result.execute()
       return("Calender Event: Succeeded in creating event.\n")
       # print("created event")
       # print("id: ", event_result['id'])
       # print("summary: ", event_result['summary'])
       # print("starts at: ", event_result['start']['dateTime'])
       # print("ends at: ", event_result['end']['dateTime'])
    except:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        exceptMessage=repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
        return("Calender Event: "+exceptMessage +"\n")


now=datetime.now()

dt_target_date=now.date()

calendar_event_status=create_calendar_event(dt_target_date, 13, "OAuthTest")

print(calendar_event_status)










        
