import requests
import os
import time
import numpy as np
import win32com.client as wincl

# Configure dates to search. Only counts business days non-holiday/break
# API call url taken from XHR response
start_date = '2020-07-06'
end_date = '2020-08-12'
date_url = f'https://calendly.com/api/booking/event_types/EEAPAX3DZQ2LU2G7/calendar/range?timezone=America%2FLos_Angeles&diagnostics=false&range_start={start_date}&range_end={end_date}&single_use_link_uuid='


r = requests.get(date_url)
response = r.json()

# Business days being returned (doesn't account for break)
# Excludes last date
days = np.busday_count(start_date, end_date)

for i in range(0, days + 1):
  if response['days'][i]['status'] == 'available':
    date = response['days'][i]['date']
    print(f"{date} has available slots")

    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak(f"{date} has available slots")
  
print('No additional dates available in specified range')


# Run the batch.bat file to loop through the script.