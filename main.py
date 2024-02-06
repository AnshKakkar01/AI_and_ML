import os
import pyaudio
import speech_recognition as sr
import pyjokes
import webbrowser
import openai
import datetime
import wikipedia
# def say(text):
#     os.system(f"say{text}")

import win32com.client

speaker=win32com.client.Dispatch("SAPI.SpVoice")


print("Enter the word you want to speak it out by computer")
s=input()
speaker.Speak(s)


def takeCommand():
        r=sr.Recognizer()
        with sr.Microphone() as source:
            # r.pause_threshold =1
            audio=r.listen(source)
            try:
                print("Recognising")
                query=r.recognize_google(audio,language="en-in")
                print(f"User said:{query}")
                return query
            except Exception as e:
                return "Some error Occured Sorry from jarviz"

if __name__ == '__main__':
    print('PyCharm')
# say("Hello i am jarviz A.I")

while True:
 print("listening...")
 query = takeCommand()
 sites=[["youtube","https://www.youtube.com"],["wikipedia","https://www.wikipedia.com"],["google","https://google.com"]]
 for site in sites:
    if f"Open {site[0]}".lower() in query.lower():
     speaker.Speak(f"Opening {site[0]} sir...")
     webbrowser.open(site[1])
    if"open music" in query:
         music = "C:\\Users\\anshk\\Downloads\\music_dir"
         songs = os.listdir(music)
         print(songs)
         os.startfile(os.path.join(music,songs[0]))
    if"stop" in query:
        break
    if "the time" in query:
                 strfTime=datetime.datetime.now().strftime("%H:%M:%S")
                 speaker.Speak(f"Sir the time is {strfTime}")
                 break
    if"joke" in query:
                        joke_1=pyjokes.get_joke(language="en",category="neutral")
                        print(joke_1)
                        speaker.Speak(joke_1)
                        break
    if"search" in query.lower():
             speaker.Speak("searching wikipedia...")
             query = query.replace("wikipedia","")
             results= wikipedia.summary(query,sentences=2)
             speaker.Speak("According to wikipedia")
             print(results)
             speaker.Speak(results)
    if"open vs code" in query:
            codepath ="C:\\Users\\anshk\\Downloads\\Microsoft VS Code\\Code.exe"
            os.startfile(codepath)
    if"stop" in query:
            speaker.Speak("Thank you")
            break



    # speaker.Speak(query)
