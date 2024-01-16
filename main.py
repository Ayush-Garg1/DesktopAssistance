import speech_recognition as sr
import os
import win32com.client
import webbrowser
import openai
import pyaudio
import datetime
import pywhatkit


def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.speak(text)


def takeCommand():
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Silence Please")
        recognizer.adjust_for_ambient_noise(source, duration=5)
        print("Speak Now...")

        try:
            audio = recognizer.listen(source, timeout=5)
            text = recognizer.recognize_google(audio)
            text = text.lower()
            print("Did you say: " + text)
            return text
        except sr.WaitTimeoutError:
            print("Listening timed out. Please speak again.")
        except sr.UnknownValueError:
            print("Speech Recognition could not understand the audio")
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")


def process(query):
    if ("bye jarvis".lower() in query.lower()) or ("by jarvis".lower() in query.lower()):
        say("bye")
        exit()
    sites =[
        ["youtube", "https://www.youtube.com"],
        ["google", "https://www.google.com"],
        ["instagram", "https://www.instagram.com"], # 'webdriver' package for opening and logging in
        ["wikipedia", "https://www.wikipedia.com"]
    ]
    # todo : add feature to play specific songs
    # todo : add feature for apps
    for site in sites:
        if f"open {site[0]}".lower() in query.lower():
            say(f"Openning {site[0]}")
            webbrowser.open(site[1])
    if "time".lower() in query.lower():
        hrs = datetime.datetime.now().strftime("%H")
        min = datetime.datetime.now().strftime("%M")
        say(f"It is {hrs} {min}")
    if "show me".lower() in query.lower():
        if "youtube" in query.lower():
            say("Playing a youtube video for this...")
            pywhatkit.playonyt(query)
            exit()
        else:
            say("Showing the the results for google search")
            pywhatkit.search(query)
            exit()


if __name__ == '__main__':
    say("Hello I'm Jarvis A I")
    while True:
        print("Listening...")
        query = takeCommand()
        # say(query)
        process(query)