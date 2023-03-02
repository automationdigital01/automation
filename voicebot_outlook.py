import streamlit as st 
import win32com.client
import datetime
import pytz
import pandas as pd

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)
events = calendar.Items
events.Sort("[Start]")
events.IncludeRecurrences = "True"

start_time = 9
end_time = 18
interval = 30
timezone = pytz.timezone("Asia/Kolkata")

slots = []
for i in range(7):
    day_start = datetime.datetime.now(timezone).replace(hour=start_time, minute=0, second=0, microsecond=0) + datetime.timedelta(days=i)
    day_end = datetime.datetime.now(timezone).replace(hour=end_time, minute=0, second=0, microsecond=0) + datetime.timedelta(days=i)
    for j in range(int((end_time - start_time) * 60 / interval)):
        slot_start = day_start + datetime.timedelta(minutes=j * interval)
        slot_end = slot_start + datetime.timedelta(minutes=interval)
        slots.append((slot_start.astimezone(pytz.utc), slot_end.astimezone(pytz.utc)))

free_slots = []
for slot in slots:
    is_free = True
    for event in events:
        if event.Start <= slot[0] and event.End >= slot[1]:
            is_free = False
            break
    if is_free:
        free_slots.append((slot[0].astimezone(timezone), slot[1].astimezone(timezone)))

free_slots_df = pd.DataFrame(free_slots, columns=['Start', 'End'])
free_slots_df['Date'] = free_slots_df['Start'].dt.date
free_slots_df['Start Time'] = free_slots_df['Start'].dt.time
free_slots_df['End Time'] = free_slots_df['End'].dt.time
free_slots_df = free_slots_df[['Date', 'Start Time', 'End Time']]
#free_slots_df.to_excel("free_slots.xlsx", index=False)
print(free_slots_df)
