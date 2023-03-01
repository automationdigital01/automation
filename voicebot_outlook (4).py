# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 12:14:28 2023

@author: gpandey2
"""

import random
import datetime
import os
import pyttsx3
import wikipedia
import speech_recognition as sr
import shutil
import win32com.client
import time
import subprocess
import pytz
import pandas as pd
import sys

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
volume = engine.getProperty('volume')
engine.setProperty('volume', 10.0)
rate = engine.getProperty('rate')
engine.setProperty('rate', rate - 25)

def speak(audio):                                # function for assistant to speak
    engine.say(audio)
    engine.runAndWait() 

def takecommand():                               # function to take an audio input from the user
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source,duration=1)
        print('Listening...')
        r.pause_threshold = 2
        audio = r.listen(source)


    try:                                            # error handling
        print('Recognizing...')
        query = r.recognize_google(audio,language = 'en-in')  # using google for voice recognition
        print(f'User said: {query}\n')

    except Exception as e :
        print('Say that again please...')        # 'say that again' will be printed in case of improper voice
        return 'None'  
    return query


def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning !")
  
    elif hour>= 12 and hour<18:
        speak("Good Afternoon !")  
  
    else:
        speak("Good Evening !") 
  
    asstname =("Christine")
    speak("I am your Assistant")
    speak(asstname)
    
def ask():
    speak("What can I do for you?")
    takecommand()
    
  
    
    

        
def username():
    speak("What should i call you sir")
    uname = takecommand()
    speak("Welcome")
    speak(uname)
    columns = shutil.get_terminal_size().columns
     
    print("#####################".center(columns))
    print("Welcome", uname.center(columns))
    print("#####################".center(columns))
     
    speak("How can i Help you")        
    
    
def free_slot():
    
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

def tomorrows_slots():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    events = calendar.Items
    events.Sort("[Start]")
    events.IncludeRecurrences = "True"
    
    start_time = 9
    end_time = 18
    interval = 30
    timezone = pytz.timezone("Asia/Kolkata")
    
    tomorrow = datetime.date.today() + datetime.timedelta(days=1)
    print(tomorrow)  
    slots=[]
    day_start = datetime.datetime.now(timezone).replace(hour=start_time, minute=0, second=0, microsecond=0) + datetime.timedelta(days=1)
    day_end = datetime.datetime.now(timezone).replace(hour=end_time, minute=0, second=0, microsecond=0) + datetime.timedelta(days=1)
    for j in range(int((end_time - start_time) * 60 / interval)):
        slot_start = day_start + datetime.timedelta(minutes=j * interval)
        slot_end = slot_start + datetime.timedelta(minutes=interval)
        slots.append((slot_start.astimezone(pytz.utc), slot_end.astimezone(pytz.utc)))
    
    tomorrows_slots = []
    for slot in slots:
        is_free = True
        for event in events:
            if event.Start <= slot[0] and event.End >= slot[1]:
                is_free = False
                break
        if is_free:
            tomorrows_slots.append((slot[0].astimezone(timezone), slot[1].astimezone(timezone)))

    tomorrows_slots_df = pd.DataFrame(tomorrows_slots, columns=['Start', 'End'])
    tomorrows_slots_df['Date'] = tomorrows_slots_df['Start'].dt.date
    tomorrows_slots_df['Start Time'] = tomorrows_slots_df['Start'].dt.time
    tomorrows_slots_df['End Time'] = tomorrows_slots_df['End'].dt.time
    tomorrows_slots_df = tomorrows_slots_df[['Date', 'Start Time', 'End Time']]
    #tomorrows_slots_df.to_excel("free_slots.xlsx", index=False)
    print(tomorrows_slots_df)

# execution control

if __name__ == '__main__' :
    clear = lambda: os.system('cls')
   
    greetings = ['hey there', 'hello', 'hi', 'Hai', 'hey!', 'hey']    
    
    clear()
    wishme()
    ask()
    
    while True:
        query = takecommand().lower()  # converts user asked query into lower case
        
        
        if query in greetings:
            random_greeting = random.choice(greetings)
            print(random_greeting)
            speak(random_greeting)
            
        
        elif 'week' in query:
            speak('getting free slots from your calendar')
            start= time.time()
            free_slot()
            end=time.time()
            print("The time of execution of above program is :", (end-start) * 10**3, "ms")
            
        elif "tomorrow's" in query:
            speak('getting tomorrows free slots from your calendar')
            tomorrows_slots()
            
            
        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f" the time is {strTime}")    
            
            
        elif "thank you" in query:
            speak("you are welcome")
            speak("anything else I can do for you")
            
            
        elif "okay bye" in query or "no bye" in query:
            speak("bye, it was nice talking to you")
            time.sleep(5)
            sys.exit("logging off") 
            
        
                 
        elif "exit" in query or "sleep" in query:
            speak("signing off")
            sys.exit("logging off")
     
        elif "log off" in query or "sign out" in query:
            speak("bye, it was nice talking to you")
            time.sleep(5)
            sys.exit("logging off")    
             

    