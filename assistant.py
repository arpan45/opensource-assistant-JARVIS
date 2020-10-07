import os
os.system('color 3f')
import speech_recognition as sr
from time import ctime
import time
from gtts import gTTS
import webbrowser
import math
from random import *
import wikipedia
import pygame
import pyautogui
pyautogui.FAILSAFE = False
import subprocess

def say(audioString):
    print(audioString)
    tts = gTTS(text=audioString, lang='en')
    tts.save("output.mp3")
    os.system("start output.mp3")
    pygame.mixer.init()
    pygame.mixer.music.load("output.mp3")
    pygame.mixer.music.play()



def recordAudio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("")
        audio = r.listen(source)


    data = ""
    try:
        data = r.recognize_google(audio)
        print("You said: " + data)
    except sr.UnknownValueError:
        print("Sorry Sir! I am unable to hear you clearly")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))

    return data

def assistant(data):
    if "lock my PC" in data:
        os.system("rundll32.exe user32.dll,LockWorkStation")
    if "put my laptop in sleep mode" in data:
        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    if "minimise windows" in data:
        os.system('''powershell -command "(new-object -com shell.application).minimizeall()"''')
    if "task view" in data :
        pyautogui.keyDown("win")
        pyautogui.press("tab")
        pyautogui.keyUp("win")
    if "close current window" in data :
        pyautogui.keyDown("alt")
        pyautogui.press("f4")
        pyautogui.keyUp("alt")
    if "show start menu" in data :
        pyautogui.press("win")

    if "type" in data :
        data = data.split(" ")
        length=len(data)
        term=data[1:length]
        pyautogui.typewrite("\t"+' '.join(term))
    if "take screenshot" in data :
        pyautogui.screenshot('screenshot.png')
    if "press enter" in data:
        pyautogui.press("enter")    
    
    if "how are you" in data:
        say("I am fine sir")
    if "hey Jarvis"in data or "hello Jarvis"in data:
        rand=randint(1,3)
        if rand==1:
            say("I am at your service Sir")
        if rand==2:
            say("Ask me anything")
        if rand==3:
            say("Yes sir! I am Here")

    if "who are you" in data:
        say("I am Jarvis. Your Personal Assistant.")

    if "what time is it" in data:
        say(ctime())

    if "open Chrome" in data:
        os.startfile(r'''C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk''')

    if "check my internet connection"in data or "check internet connection" in data:
        hostname="google.co.in"
        response=os.system("ping -c 1"+hostname)
        if response==0:
            say("I Think Internet is Disconnected")
        else:
            say("Internet Connection is fine Sir")

    #if "sing a song Jarvis" in data:
        #os.system("jarvis_song.mp3")

    if "Google search" in data:
        data = data.split(" ")
        length=len(data)
        if length>3:
            #for i in range(2,length-1):
            term=data[2:length]

        elif length==3:
            term=data[2]
        else :
            term=""

        url = "https://www.google.co.in/search?q={}".format(' '.join(term))
        webbrowser.open_new_tab(url)

    if "calculate" in data:
        data = data.split(" ")
        if data[2]=="+":
            num1=data[1]
            num2=data[3]
            add=int(num1)+int(num2)
            str_add=str(add)
            say("The Required answer is"+str_add)
        if data[2]=="multiply":
            num1=data[1]
            num2=data[3]
            add=float(num1)*float(num2)
            str_add=str(add)
            say("The Required answer is " + str_add)
        if data[1]=="factorial":
            fac=math.factorial(int(data[2]))
            str_fac=str(fac)
            say(str_fac)

    if "take a note" in data:
        data=data.split(" ")
        file=open("note.txt", "a")
        file.write("\n"+' '.join(data))
        file.close()
        say("Note Taken Sir. Any thing Else?")
        data = recordAudio()
        if "no" in data:
            say("Ok.Sir!")
        if "yes" in data:
            say("Go on Sir")

    if "who is" in data or "what is" in data:
        data=data.split(" ")
        length=len(data)
        if length>3:
            term=data[2:length]

        elif length==3:
            term=data[2]
        else :
            term=""
        ans=wikipedia.summary(term, sentences=2)
        say(ans)

    if "where is" in data:
        data = data.split(" ")
        location = data[2]
        say("Just A Second Sir, I will show you where " + location + " is.")

        URL = "https://www.google.com/maps/place/" + location + "/&amp;"
        webbrowser.open(URL, new=2)

    if "open Explorer" in data:
        os.startfile(r'''C:\Windows\EXPLORER.EXE''')
        
    if "open my files" in data:
        username = os.environ.get('USERNAME')
        subprocess.Popen(f'C:\Windows\EXPLORER.EXE /select,C:\\Users\\{username}\\Documents\\')
        
time.sleep(2)
say("Hello Sir, what can I do for you?")
while 1:
    data = recordAudio()
    assistant(data)
