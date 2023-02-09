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
slots = []
timezone = pytz.timezone("Asia/Kolkata")

for i in range(7):
    day_start = (datetime.datetime.now() + datetime.timedelta(days=i)).replace(hour=start_time, minute=0, second=0, microsecond=0)
    day_end = (datetime.datetime.now() + datetime.timedelta(days=i)).replace(hour=end_time, minute=0, second=0, microsecond=0)
    for j in range(int((end_time - start_time) * 60 / interval)):
        slot_start = day_start + datetime.timedelta(minutes=j * interval)
        slot_end = slot_start + datetime.timedelta(minutes=interval)
        slots.append((timezone.localize(slot_start), timezone.localize(slot_end)))
   


free_slots = []
for slot in slots:
    is_free = True
    for event in events:
        if event.Start <= slot[0] and event.End >= slot[1]:
            is_free = False
            break
    if is_free:
        free_slots.append(slot)
        free_slots_df = pd.DataFrame(free_slots, columns=['Start', 'End'])
        free_slots_df['Start'] = free_slots_df['Start'].dt.tz_localize(None)
        free_slots_df['End'] = free_slots_df['End'].dt.tz_localize(None)

        free_slots_df.to_excel("free_slots.xlsx", index=False)
        


    
    


    



    
