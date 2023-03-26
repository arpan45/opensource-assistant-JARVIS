from subprocess import getoutput
import json
import pyttsx3
import pyaudio
from vosk import Model, KaldiRecognizer
import pyautogui
import wikipedia
from random import *
import webbrowser
from gtts import gTTS
import time
from time import ctime
import os
from getpass import getuser
import shutil
import win10toast
from pygame import mixer
import pyperclip




roll = [
    "You’re welcome",
    "don’t mention it",
    "My pleasure",
    "it’s nothing",
    "Not at all",
    "that’s the least I could do"
    ]

music_path = None
music_list = []

def notif(title, msg, dur=2):
    notif = win10toast.ToastNotifier()
    notif.show_toast(title,msg,None,dur)

pyautogui.FAILSAFE = False


model = Model(r"\vosk\en-small")

recognizeren = KaldiRecognizer(model, 48000)
cap = pyaudio.PyAudio()
stream = cap.open(format=pyaudio.paInt16, channels=1,
                  rate=48000, input=True, frames_per_buffer=4096)
stream.start_stream()

def printf(text):
    print(text.center(shutil.get_terminal_size().columns))

notif("Assistant","API is loaded ...")

try:
    app = json.load(open("app.json", "r"))
    print("LOG (JsonLib:loadfile) app.json Done")
except:
    print("LOG (JsonLib:loadfile) app.json Filed")



def say(audioString):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty("rate", 120)
    # engine.setProperty("voice",voices[0].id)  # male voice
    engine.setProperty("voice", voices[1].id)  # female voice
    engine.say(audioString)
    engine.runAndWait()
    engine.stop()

notif("Assistant","I can listen to what you are saying")
say("I can listen to what you are saying")

def recordAudio():
    result = None
    data = stream.read(8192)
    if recognizeren.AcceptWaveform(data):
        result = recognizeren.Result()
        result = json.loads(result)
        result = result["text"]
        return result

os.system("cls")

def assistant(data):
    if data == None:
        pass
    else:
        print(data)
    if data != None:
        if "lock" in data and "pc" in data:
            printf(data)
            os.system("rundll32.exe user32.dll,LockWorkStation")
            say("Your system has been locked")
        elif "say" in data and "my" in data and "name" in data:
            printf(data)
            say(getuser())
        elif "sleep" in data and "pc" in data:
            printf(data)
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        elif "minimise" in data or "minimize" in data and "windows" in data:
            printf(data)
            os.system(
                '''powershell -command "(new-object -com shell.application).minimizeall()"''')
        elif "task" in data and "view" in data:
            printf(data)
            pyautogui.keyDown("win")
            pyautogui.press("tab")
            pyautogui.keyUp("win")
        elif "close" in data and "current" in data and "window" in data:
            printf(data)
            pyautogui.keyDown("alt")
            pyautogui.press("f4")
            pyautogui.keyUp("alt")
        elif "show" in data and "start" in data and "menu" in data:
            printf(data)
            pyautogui.press("win")
        elif "show" in data and "run" in data :
            printf(data)
            pyautogui.keyDown("alt")
            pyautogui.press("space")
            pyautogui.keyUp("alt")
        elif "type" in data:
            printf(data)
            data = data.split(" ")
            length = len(data)
            term = data[1:length]
            pyautogui.typewrite("\t"+' '.join(term))
        elif "take screenshot" in data:
            printf(data)
            pyautogui.screenshot('screenshot.png')
        elif "press enter" in data:
            printf(data)
            pyautogui.press("enter")

        elif "what time is it" in data:
            printf(data)
            say(ctime())

    
        elif "check" in data and "internet" in data or "internet" in data and "connection" in data:
            printf(data)
            hostname = "google.co.in"
            response = os.system("powershell ping "+hostname)
            if response == 0:
                say("I Think Internet is Disconnected")
            else:
                say("Internet Connection is fine Sir")

        
        elif "search google" in data:
            printf(data)
            data = data.split(" ")
            length = len(data)
            if length > 3:
                # for i in range(2,length-1):
                term = data[2:length]

            elif length == 3:
                term = data[2]
            else:
                term = ""

            url = "https://www.google.co.in/search?q={}".format(' '.join(term))
            webbrowser.open_new_tab(url)

        elif "find movie" in data:
            printf(data)
            data = data.split(" ")
            length = len(data)
            if length > 3:
                # for i in range(2,length-1):
                term = data[2:length]

            elif length == 3:
                term = data[2]
            else:
                term = ""

            url = "https://www.imdb.com/find/?q={}".format(' '.join(term))
            webbrowser.open_new_tab(url)

        elif "who is" in data or "what is" in data:
            try:
                printf(data)
                data = data.split(" ")
                length = len(data)
                if length > 3:
                    term = data[2:length]

                elif length == 3:
                    term = data[2]
                else:
                    term = ""
                ans = wikipedia.summary(term, sentences=1)
                say(ans)
            except wikipedia.exceptions.DisambiguationError:
                say("Can you tell me exactly again?")

        elif "where" in data and "is" in data:
            printf(data)
            data = data.split(" ")
            location = data[2]
            say("Just A Second Sir, I will show you where " + location + " is.")

            URL = "https://www.google.com/maps/place/" + location + "/&amp;"
            webbrowser.open(URL, new=2)

        elif "open" in data:
            printf(data)
            if "game" in data and "loop" in data:
                os.startfile(app["gameloop"])

        elif "thank" in data and "you" in data or "thanks" in data and "honey" in data:
            printf(data)
            say(choice(roll))

        elif "who are you" == data:
            printf(data)
            say("I am your dear")

        elif "shut" in data and "up" in data or "fuck" in data and "you":
            printf(data)
            if "shut" in data:
                say("hey hey\nshut up yourself")
            elif "fuck" in data:
                say("hey hey\nDamn you, I didn't do anything")

        elif data == "honey":
            printf(data)
            say("Yes Dear?")

        elif "good" in data and "by" in data:
            say("by by dear")
            exit()

        elif "start" in data and "music" in data and "engine" in data:
            mixer.init()
            music_path = pyperclip.paste()
            os.chdir(music_path)
            m = getoutput(f"dir /s /b")

            m = m.split("\n")
            for file in m:
                if file[-4:] == ".mp3":
                    music_list.append(file)
                else:
                    pass
            global set_item
            set_item = 1
            say("run engine")
            


            

        elif "play" in data and "music" in data:
            # try:
                say("ok")
                print(music_list[set_item])
                mixer.music.play(str(music_list[set_item]))
            # except:
            #     say("Can you copy a music folder and then say 'run the music engine'? I can't do this alone")

        elif "pause" in data and "music" in data:
            mixer.music.pause()

        elif "stop" in data and "music" in data:
            mixer.music.stop()

        elif "next" in data and "music" in data:
            
            mixer.music.stop()
            if len(music_list) <= set_item:
                set_item = set_item + 1
            else:
                set_item = 0
            mixer.music.load(music_list[set_item])
            mixer.music.play()


time.sleep(2)

while True:
    data = recordAudio()
    assistant(data)
