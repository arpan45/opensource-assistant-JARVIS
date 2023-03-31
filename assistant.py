import json
import pyaudio
from vosk import Model, KaldiRecognizer
import pyautogui
import wikipedia
import webbrowser
import time
import os
import shutil
import pyttsx3
from getpass import getuser
import requests

music_path = None
music_list = []

pyautogui.FAILSAFE = False

model = Model(r"\vosk\en-small")

recognizeren = KaldiRecognizer(model, 48000)
cap = pyaudio.PyAudio()
stream = cap.open(format=pyaudio.paInt16, channels=1,
                  rate=48000, input=True, frames_per_buffer=4096)
stream.start_stream()


def printf(text):
    print(text.center(shutil.get_terminal_size().columns))


def recordaudio():
    result = None
    stream_value = stream.read(8192)
    if recognizeren.AcceptWaveform(stream_value):
        result = recognizeren.Result()
        result = json.loads(result)
        result = result["text"]
        return result


def say(text):
    engine = pyttsx3.init()
    engine.setProperty('rate', 120)
    engine.setProperty('volume', 1.0)
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.say(text)
    engine.runAndWait()


os.system("cls")


def assistant(value):
    if value is not None:
        print(value)
        if "lock" in value and "pc" in value:
            say("i will lock the device right now")
            os.system("rundll32.exe user32.dll,LockWorkStation")
        elif "say" in value and "my" in value and "name" in value:
            say("your name is " + getuser())
        elif "sleep" in value and "pc" in value:
            say("i will put the device to sleep right now")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        elif "minimise" in value or "minimize" in value and "windows" in value:
            os.system(
                '''powershell -command "(new-object -com shell.application).minimizeall()"''')
        elif "task" in value and "view" in value:
            say("ok. run task view")
            pyautogui.keyDown("win")
            pyautogui.press("tab")
            pyautogui.keyUp("win")
        elif "close" in value and "current" in value and "window" in value:
            pyautogui.keyDown("alt")
            pyautogui.press("f4")
            pyautogui.keyUp("alt")
        elif "type" in value:
            value = value.split(" ")
            length = len(value)
            term = value[1:length]
            pyautogui.typewrite("\t" + ' '.join(term))
        elif "take screenshot" in value:
            say("1, 2, 3, i take now screenshot")
            pyautogui.screenshot('screenshot.png')
        elif "press enter" in value:
            pyautogui.press("enter")
        elif "what time is it" in value:
            say(time.ctime())
        elif "check" in value and "internet" in value or "internet" in value and "connection" in value:
            try:
                _ = requests.head('https://www.google.com/', timeout=5)
                say("i think you are connected")
            except requests.ConnectionError:
                say("i think you are not connected")
        elif "search google" in value:
            value = value.split(" ")
            length = len(value)
            if length > 3:
                # for i in range(2,length-1):
                term = value[2:length]

            elif length == 3:
                term = value[2]
                say("i am doing a search about " + term)
            else:
                term = ""
                say("i did not understand what happened")

            url = "https://www.google.co.in/search?q={}".format(' '.join(term))
            webbrowser.open_new_tab(url)
        elif "who is" in value or "what is" in value:
            try:
                say("wait a while")
                value = value.split(" ")
                length = len(value)
                if length > 3:
                    term = value[2:length]

                elif length == 3:
                    term = value[2]
                else:
                    term = ""
                ans = wikipedia.summary(term, sentences=1)
                say(ans)
            except wikipedia.exceptions.DisambiguationError:
                say("can you tell me exactly again?")
        elif "where" in value and "is" in value:
            value = value.split(" ")
            location = value[2]
            say(f"it is very good. let's go see where {location} is")
            url = "https://www.google.com/maps/place/" + location + "/&amp;"
            webbrowser.open(url, new=2)
        elif "thank" in value and "you" in value or "thanks" in value and "honey" in value:
            say("thankful")
        elif "who are you" == value:
            say("i am a voice assistant")
        elif "shut" in value and "up" in value or "fuck" in value and "you":
            printf(value)
            if "shut" in value:
                say("shut up yourself")
            elif "fuck" in value:
                say("impolite!")
        elif "good" in value and "by" in value:
            say("by by dear")
            exit()
        elif "kill" in value and "yourself" in value:
            if "ok" in value:
                say("ok. by")
                exit()
            else:
                say("do you really want this?")
        elif "chenge" in value:
            if "desktop" in value:
                pyautogui.keyDown("win")
                pyautogui.press("tab")
                pyautogui.keyUp("win")
            elif "app" in value:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(3)
                pyautogui.keyUp("alt")
        elif "show" in value:
            if "desktop" in value:
                pyautogui.keyDown("win")
                pyautogui.press("d")
                pyautogui.keyUp("win")
            elif "widget" in value:
                pyautogui.keyDown("win")
                pyautogui.press("w")
                pyautogui.keyUp("win")
            elif "explorer" in value:
                pyautogui.keyDown("win")
                pyautogui.press("e")
                pyautogui.keyUp("win")
            elif "start" in value and "menu" in value:
                pyautogui.press("win")
            elif "run" in value:
                say("ok. this is run")
                pyautogui.keyDown("alt")
                pyautogui.press("space")
                pyautogui.keyUp("alt")
        elif "action" in value:
            if "copy" in value:
                pyautogui.keyDown("ctrl")
                pyautogui.press("c")
                pyautogui.keyUp("ctrl")
            elif "paste" in value:
                pyautogui.keyDown("ctrl")
                pyautogui.press("v")
                pyautogui.keyUp("ctrl")
            elif "cut" in value:
                pyautogui.keyDown("ctrl")
                pyautogui.press("x")
                pyautogui.keyUp("ctrl")
            elif "select" in value and "everything" in value:
                pyautogui.keyDown("ctrl")
                pyautogui.press("a")
                pyautogui.keyUp("ctrl")
        elif "music" in value:
            if "next" in value:
                pyautogui.press("nexttrack")
            elif "back" in value:
                pyautogui.press("nexttrack")
            elif "play" in value or "pause" in value:
                pyautogui.press("playpause")

        # elif "start" in data and "music" in data and "engine" in data:
        #     mixer.init()
        #     path = pyperclip.paste()
        #     os.chdir(path)
        #     m = getoutput(f"dir /s /b")
        #
        #     m = m.split("\n")
        #     for file in m:
        #         if file[-4:] == ".mp3":
        #             music_list.append(file)
        #         else:
        #             pass
        #     global set_item
        #     set_item = 1
        #     say("run engine")

        # elif "play" in value and "music" in value:
        #     # try:
        #     # say("ok")
        #     print(music_list[set_item])
        #     mixer.music.play(str(music_list[set_item]))
        # # except:
        # #     say("Can you copy a music folder and then say 'run the music engine'? I can't do this alone")
        #
        # elif "pause" in value and "music" in value:
        #     mixer.music.pause()
        #
        # elif "stop" in value and "music" in value:
        #     mixer.music.stop()
        #
        # elif "next" in value and "music" in value:
        #
        #     mixer.music.stop()
        #     if len(music_list) <= set_item:
        #         set_item = set_item + 1
        #     else:
        #         set_item = 0
        #     mixer.music.load(music_list[set_item])
        #     mixer.music.play()


time.sleep(2)

while True:
    data = recordaudio()
    assistant(data)
