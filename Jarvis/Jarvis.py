"""
What is pyttsx3?
A python library which will help us to convert text to speech. In short, it is a text-to-speech library.
It works offline, and it is compatible with Python 2 as well the Python 3.

What is sapi5?
Speech API developed by Microsoft.
Helps in synthesis and recognition of voice

What Is VoiceId?
Voice id helps us to select different voices.
voice[0].id = Male voice
voice[1].id = Female voice
"""

import pyttsx3
import wikipedia
import datetime
import speech_recognition as sr
import webbrowser
import os
import smtplib
import random
import wolframalpha
import json
import requests

f = open("input.txt")   #make a file to safe your password in that file
inp=f.read()            #read the password
# print(inp)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')    #getting details of current voice
# print(voices[0].id)
engine.setProperty('voice',voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()     #Without this command, speech will not be audible to us.

# def speak(str):
#     from win32com.client import Dispatch
#     speak = Dispatch("SAPI.SpVoice")
#     speak.Speak(str)

def wishMe():
    '''Here, we have stored the integer value of the current hour or time into a variable named hour.Now, we will use this
    hour value inside an if - else loop.'''
    hour = (datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Sir!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir!")
    else:
        speak("Good Evening Sir!")
    speak("I am Jarvis. Please tell me how may I help you?")

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1   # seconds of non-speaking audio before a phrase is considered complete
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')  # Using google for voice recognition.
        print(f"User said: {query}\n")  # User query will be printed.

    except Exception as e:
        # print(e)
        speak("Could not understand your audio, PLease try again !")  # Say that again will be printed in case of improper voice
        return "None"  # None string will be returned
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('Your email address', inp)     #inp will fetch the password from the input file.
    server.sendmail('Your email address', to, content)
    server.close()

if __name__ == '__main__':
    # speak("Hello Everyone")
    wishMe()
    while True:
        query = takeCommand().lower()   #Converting user query into lower case

        # Logic for executing tasks based on query
        if 'wikipedia' in query:     #if wikipedia found in the query then this block will be executed
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query,sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir?")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Kshitij.")

        elif 'find answer' in query:
            # Taking input from user
            question = input('Question: ')
            # App id obtained by the above steps
            app_id = 'JHJ5UK-X5T22JRYGG'
            # Instance of wolf ram alpha
            # client class
            client = wolframalpha.Client(app_id)
            # Stores the response from
            # wolframalpha
            res = client.query(question)
            # Includes only text from the response
            answer = next(res.results).text
            print(answer)
            speak(f"The answer is {answer}")

        elif 'thank you jarvis' in query:
            speak("Your Welcome Sir!")

        elif 'open youtube' in query:
            speak("Opening Youtube...")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Opening Google...")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Opening Stackoverflow...")
            webbrowser.open("stackoverflow.com")

        elif "open chrome" in query:
            speak("Google Chrome")
            os.startfile('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')

        elif 'play music' in query or 'play songs' in query:
            music_dir = 'D:\\Songs'  #\\ to escape character
            songs = os.listdir(music_dir)
            # print(songs)
            l=len(songs)
            select=songs[random.randint(0,l-1)]
            os.startfile(os.path.join(music_dir,select))

        elif 'read news' in query:
            speak("News for today.. Lets begin")
            url = "https://newsapi.org/v2/top-headlines?sources=the-times-of-india&apiKey=38ad506c3e7d4c15996228b61194edda"  #generate your own api key on newsapi.org
            news = requests.get(url).text
            news_dict = json.loads(news)
            arts = news_dict['articles']
            for article in arts:
                print(article['title'])
                speak(article['title'])
                speak("Would you like to listen to next news?")
                command=takeCommand()
                if 'yes' in command or 'sure' in command:
                    speak("Moving on to the next news..Listen Carefully")
                    continue
                elif 'no' in command:
                    speak("Thanks for listening...")
                    break
            # speak("Thanks for listening...")

        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            print(strTime)
            speak(f"The time is {strTime}")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("Whom should i send?")
                print("Enter the email-id:")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry Sir, I am not able to send this email")


        elif 'quit' in query:
            speak("Good bye Sir..")
            exit()





