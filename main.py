import win32com.client
import speech_recognition as sr
import webbrowser
import os
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        print("Listening...")
        speaker.Speak("Listening...")
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-IN")
            print(f"User said: {query}")
            return query
        except Exception as e:
            print("Some Error occurred! Sorry from Jarvis.")
            return None

if __name__ == '__main__':
    speaker.Speak("Hello! I am Jarvis AI")
    query = takeCommand()
    #TODO: ADD MORE SITES TO VIEW
    sites=[["youtube", "https://youtube.com"],["wikipedia","https://wikipedia.com"],
           ["google","https://google.com"]]
    for site in sites:
        if f"Open {site[0]}".lower() in query.lower():
            speaker.Speak(f"Opening {site[0]}, sir...")
            webbrowser.open(site[1])
    #TODO: ADD FEATURE FOR MUSIC
    if "open music"in query:
        musicPath = r"C:\Users\Saad Zubair\Downloads\Egzod, Maestro Chives, Neoni - Royalty [NCS Release].mp3"
        speaker.Speak("Playing Music, sir...")
        os.startfile(musicPath)
    
    if "the time"in query:
        strfTime=datetime.datetime.now().strftime("%H:%M:%S")
        speaker.Speak(f"Sir, the time is {strfTime}")
        
    if "open code studio" in query.lower():
        speaker.Speak("Opening Microsoft Visual Studio 2022, sir...")
        vs_path = r"C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"  # Update this to your actual Visual Studio path
        os.startfile(vs_path)

    
