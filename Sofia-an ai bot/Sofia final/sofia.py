import sys
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import os
import instaloader
import cv2
import webbrowser
import random
import requests
from requests import get
import pywhatkit as kit
import PyPDF2
import pyjokes
import operator #for calucation using voice
from pywikihow import search_wikihow
from bs4 import BeautifulSoup
import pyautogui
import psutil
import subprocess
#import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import ssl, smtplib
#smtplib is the module used by python to send emails through SMTP.
#The ssl module is used to access the Operating System Socket.
import webbrowser
import os
#import winshell
import pyjokes
#import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
#from client.textui import progress
# from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import speedtest
import time
from twilio.rest import Client
from PyQt5 import QtWidgets ,QtCore, QtGui
from PyQt5.QtCore import QTimer ,QTime ,QDate ,Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from geopy.distance import great_circle
from geopy.geocoders import Nominatim
import geocoder
import googlemaps
from playsound import playsound
# import winsound #pip install Playsound
import keyboard
from PyDictionary import PyDictionary as Diction
from tkinter import Label
from tkinter import Entry
from tkinter import Button
from tkinter import Tk
from tkinter import StringVar
from pytube import YouTube
from googletrans import Translator
#youtube automation #google automation


engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
#printvoices
print(voices[1].id)
engine.setProperty('voice',voices[1].id)
engine.setProperty('rate', 180)



def speak(audio):
    print("   ")
    engine.say(audio)
    print(f"AI Assistant:{audio}")
    engine.runAndWait()
    
def wishMe():    
    hour =int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning")
        
    elif hour>=12 and hour<19:
        speak("Good Afternoon!")
    else:
         speak("Good Evening") 
     
    strTime = datetime.datetime.now().strftime("%H:%M:%S")    
    speak(f"its {strTime}")
    speak("I am Sofia Sir. your Personal AI Assistant . Please Tell me , how  may i help you sir.")






#for news updates
def news(): 
   
    main_url= 'http://newsapi.org/v2/top-headlines?sources=techcrunch&apiKey=c15be20b3bd44f0fbe7866085aa119bf'
    main_page =  requests.get(main_url).json()
    #print(main page)
    articles =main_page["articles"]
    print (articles)
    head=[]
    day=["first","second","third","forth","fifth","sixth","seventh","eighth","ninth","tenth"]
    for ar in articles:
        head.append(ar["title"])
    for i in range (len(day)):
        #print ftoday's {}day[i]} news is :",head[i]}
        speak(f"today's {day[i]} news is :  {head[i]}")
# def SpeedTest():
#     speak("checking speed ...")
#     speed=speedtest.Speedtest()
#     downlaoding=speed.downlaod()
#     correctDown=int(downlaoding/800000)
#     uploading=speed.upload()
#     correctUpload=int(uploading/800000)
#     speed("the downloading is {correctDown} and Uploading Speed is {correctUpload} mbps")
        
def alarm(Timing):
    altime =str(datetime.datetime.now().strptime(Timing,"%I:%M %p"))
    
    altime = altime[11:-3]
    print(altime)
    Horeal = altime[:2]
    Horeal = int(Horeal)
    Mireal = altime[3:5]
    Mireal = int(Mireal)
    print(f"Done,alarm is set for {Timing}")
    speak(f" Done sir , alarm is set for {Timing} . till then have a break , have a kitkat sir . i will wake u up .")
    
    while True:
        if Horeal ==datetime.datetime.now().hour:
            if Mireal==datetime.datetime.now().minute:
                speak("alarm is running")
                #os.startfile('C:\\Users\\RohanRVC\\.spyder-py3\\annoying sound music.mp3')
                playsound('annoying sound music.mp3')
                speak("alarm closed sir . and i am ready for next work sir")
                
            elif Mireal<datetime.datetime.now().minute:
                break
        else:
            pass
def pdf_reader():
    cm=takeCommand().lower()
    book = open('dsc task.pdf','rb')
    pdfReader= PyPDF2.PdfFileReader(book) #pip install PyPDF2    
    pages = pdfReader.numPages
    speak(f"Total numbers of pages in this book is {pages}")
    speak("sir from which page  u want me to start reading? .")   
    #pg = int(input("please enter tyhe page number here sir:-"))
    cm= takeCommand().lower()
    page = pdfReader.getPage(int(cm))
    text = page.extractText()
    speak(text)
    speak("pdf reading done sir .")

def Googlemaps(Place):
    speak("ok sir , gathering details through satellite , please wait a second")
    Url_Place="https://www.google.com/maps/place/" + str(Place)
    geolocator =Nominatim(user_agent="mycoder")
    location =geolocator.geocode(Place , addressdetails= True)
    target_latlon = location.latitude , location.longitude
    location = location.raw['address']
    target = {'city' : location.get('city',''),
              'state'  : location.get('state',''),
               'country' : location.get('country','')}
    print(target) 
    current_loca = geocoder.ip('me')
    current_latlon = current_loca.latlng
    distance =str(great_circle(current_latlon,target_latlon))
    distance = str(distance.split(' ',1)[0])
    distance = round(float(distance),2)
    speak(target)
    webbrowser.open(url=Url_Place)
    speak(f"sir , {Place} is {distance} kilometer away from your location")
    print(speak)

def Emotion():
     speak("Hello sir how are u feeling")
     cm=takeCommand()
     if 'good' in cm:
         speak("good to hear sir")
     elif 'fine' in cm:
         speak("how fine sir?")
     elif 'not' in cm or 'negative' in cm:
         speak("i am sorry to hear that sir")
     
    
    
    
def Music(): 
    speak("Tell me the name of the song sir !")
    cm=takeCommand()
    if 'see you again' in cm:
        os.startfile('C:\\Users\\RohanRVC\\Documents\\song\\Wiz Khalifa - See You Again ft. Charlie Puth [Official Video] Furious 7 Soundtrack(M4A_128K)_1.m4a')
    elif 'ye lila' in cm :
        os.startfile('C:\\Users\\RohanRVC\\Documents\\song\\Ye Lili ye Lila  song for vipul parmar.mp3')
    else:
        kit.playonyt(cm)
# def TakeHindi():
#      #it takes micorphone input from the user and return strin g output 
#         r=sr.Recognizer()
#         with sr.Microphone() as source:
#             print("Listening...")
#             r.pause_threshold=1
#             #r.adjust_for_ambient_noice(source)
#             #audio = r.listen(source)
#             audio=r.listen(source)
            
#         try:
#              print("Recognizing...")
#              query=r.recognize_google(audio, language='en-in')
#              query=r.recognize_google(audio, language='en-in')
#              print(f"Boss said:{query}\n")
#         except Exception as e:
#             print(e)
#             #speak("say that again please...")
#             print("Say that Again Please....")
#             return"none"
#         query=query.lower()
#         return query 
# def Tran():
#     speak("Tell me the line sir")
#     line = TakeHindi()
#     traslate = Translator()
#     result = traslate.translate(line)
#     Text=result.text
#     speak(Text)

def Whatsapp():
    speak("tell me the name of the person ! whom u want to send message sir")
    cm = takeCommand()
    if "divya" in cm:
        speak("tell me the message sir that u want to send")
        msg = takeCommand()
        speak("tell me the time sir")
        speak("time in hours!")
        hour = int(takeCommand())
        #cm1= takeCommad()
        speak("time in minutes")
        min=int(takeCommand())
        #min = takeCommand()
        kit.sendwhatmsg("+917798830125",msg,hour,min,30)
        speak("ohk sir,sending whatsapp message !")
    elif "priyanshu" in cm:
        speak("tell me the message sir that u want to send")
        msg = takeCommand()
        speak("tell me the time sir")
        speak("time in hours , enter now sir!")
        hour = int(takeCommand())
        #hour = takeCommad()
        speak("time in minutes")
        min=int(takeCommand())
        #min = takeCommand()
        kit.sendwhatmsg("+918766933423",msg,hour,min,30)
        speak("ohk sir,sending whatsapp message !")
    elif "papa" in cm:
        speak("tell me the message sir that u want to send")
        msg = takeCommand()
        speak("tell me the time sir")
        speak("time in hours , enter now sir!")
        hour = int(takeCommand())
        #hour = takeCommad()
        speak("time in minutes")
        min=int(takeCommand())
        #min = takeCommand()
        kit.sendwhatmsg("+919819882492",msg,hour,min,30)
        speak("ohk sir,sending whatsapp message !")
    else:
        speak("tell me the phone number sir or enter here sir ")
        #phone=takecommand() 
        ph=int(input('enter phone no here sir-:'))
        speak("tell me the message sir that u want to send")
        msg = takeCommand()
        speak("tell me the time sir")
        speak("time in hours!")
        #hour = takeCommad()
        hour = int(takeCommand())
        speak("time in minutes")
        #min = takeCommand()
        min=int(takeCommand())
        kit.sendwhatmsg(ph,msg,hour,min,30)
        speak("ok sir,sending whatsapp message !")
    speak("message sent sir")
    
def YoutubeAuto():
        speak("sir, What should  i play on youtube ?")
        #songg = input("enter song here:-")
        cm= takeCommand().lower()
        #print(speak)
        #webbrowser.open(f"www.youtube.com/{name}")            
        
        kit.playonyt(cm)
        com =takeCommand()
        if "pause" in com:
            keyboard.press('space bar')
        elif "mute" in com:
            keyboard.press('m')
        elif "skip" in com or "forward" in com:
            keyboard.press("l")
        elif "restart" in com:
            keyboard.press("0")
        elif "back" in com or "backward" in com:
            keyboard.press("j")
        elif "full screen" in com:
            keyboard.press("f")
        elif "subtitle" in com or "caption" in com:
            keyboard.press("c")
        elif "theatre mode" in com:
            keyboard.press("t")
        elif "mini play" in com:
            keyboard.press("i")
        speak("done sir")
        
# def WebsiteAutomation():
#     we = takeCommand()
#     speak("Website automation started")
    
#     if "linkedin" in we:
#         webbrowser.open('https://www.linkedin.com/')
#     elif "movie" in we:
#         webbrowser.open('https://moviesverse.org.in/')
#     elif "article rewriter" in we:
#         webbrowser.open("https://articlerewritertool.com/")
#         speak("article rewriter opened sir")
        

        
def ChromeAuto():
        speak("chrome automation started")
        chr = takeCommand().lower()
        if "close this tab" in chr:
            keyboard.press_and_release('ctrl + w')
            speak("tab closed sir")
        elif "hompepage" in chr:
            keyboard.press_and_release("alt + home")
            speak("home page is here sir")
        elif "activate full screen mode" in chr  :
            keyboard.press("f11")
            speak("sir activated full screen mode")
        elif "close full screen mode" in chr :
            keyboard.press("f11")
            speak("deactivated full screen mode")
        elif "open new tab" in chr:
            keyboard.press_and_release("ctrl + t")
            speak("opened new tab sir")
        elif "stop loading" in chr:
            keyboard.press("esc")
            ("stoped sir")
        elif "before tab" in chr:
            keyboard.press_and_release("ctrl +1")
            speak("moved sir")
        elif "next tab" in chr:
            keyboard.press_and_release("ctrl + 8")
            speak("moved sir")
        elif "reset browser zoom" in chr or "default browser size" in chr:
            keyboard.press_and_release("ctrl + 0")
            speak("browser set to its default size")
        # elif "hide bookmark" in chr  : 
        #     keyboard.press_and_release("ctrl + shift + B")
        #     speak("bookmark hidden sir")
        # elif "show bookmark" in chr or "unhide bookmark" in chr : 
        #     keyboard.press_and_release("ctrl + shift + B")
        #     speak("bookmark is now visible sir to everyone")
        elif ("last tab") in chr:
            keyboard.press_and_release("ctrl + 9")
            speak("switched to last tab of this browser sir")
        elif "open file in browser" in chr:
            keyboard.press_and_release("ctrl + O")
            speak("opened file in browser sir")
        elif "select everything" in chr:
            keyboard.press_and_release("ctrl + A")
            speak("everything selected sir")
        elif "add this to bookmark" in chr:
            keyboard.press_and_release("ctrl + D")
            speak("added this page to bookmark sir")
        elif "bookmark manager" in chr or "book manager" in chr:
            keyboard.press_and_release("ctrl + shift + O")
            speak('opened manger of the bookmark sir')
        elif "download window" in chr or "show download" in chr:
            keyboard.press_and_release("ctrl + J")
            speak("here is the downlaod window sir")
        elif "browser history" in chr or "history" in chr:
            keyboard.press_and_release("ctrl + H")
            speak("here is the browsing history sir")
            # while True:
            #         speak("Please tell sir do you what to clear your browsing history?")
            #         mia= takeCommand().lower()
            #         try:
            #             if "exit" in mia or "close" in mia or "no" in mia:
            #                 speak("okay sir , i won't delete your browsing history sir . though i would give an suggestion sir its better to use incognito tab")
            #                 break
            #             elif "yes" in mia:
            #                 speak("ohk sir , to confirm that orders are directly coming from you , u know what to do sir , better than me sir") #press ok
            #                 keyboard.press_and_release("ctrl + shift + delete")
            #         except:
            #             speak("not able to detect thats from whom this order is coming")
        elif "new browser" in chr:
            keyboard.press_and_release("ctrl + N")
            speak("opened new browser sir")
        elif "incognito mode" in chr or "private mode" in chr:
            keyboard.press_and_release("ctrl + shift + N")
            speak("switched to private aka incognito mode . enjoying surfing now sir")
        elif "print page" in chr or "print this page" in chr:
            keyboard.press_and_release("ctrl + P")
            speak("page is printed sir")
        elif "refresh the page" in chr or "refresh this page" in chr:
            keyboard.press("f5")
            speak("page is refreshed sir")
        elif "save the page" in chr or "save this page" in chr:
            keyboard.press_and_release("ctrl + S")
            speak("to save the page press ok , . sir now page is saved ")
        elif "new tab" in chr :
            keyboard.press_and_release("ctrl + T")
            speak("opened new tab sir")
        elif "source code" in chr:
            keyboard.press_and_release("ctrl + U")
            speak("here is the source code of this page sir")
        elif "close this browser" in chr:
            keyboard.press_and_release("ctrl + shift + W")
            speak("browser has been closed sir")
        elif "reopen" in chr or "reopen tab" in chr or "reopen browser" in chr:
            keyboard.press_and_release("ctrl + shift + T")
            speak("we are back in the same tab sir")
        speak("ok sir done")
    
def Dict():
    speak("activated dictionary mode")
    speak("what  do u wanna know sir . meaning , synonym or antonym ?")
    bp = takeCommand().lower()
    
    if "meaning" in bp:
        speak("tell me the word sir")
        cp = takeCommand().lower()
        
        result=Diction.meaning(cp)
        speak(f"The Meaning of {cp} is {result}")
    if "synonym" in bp:
        speak("tell me the word sir")
        cg = takeCommand().lower()
        result=Diction.meaning(cg)
        speak(f"The Synonym of {cg} is {result}")
    if "antonym" in bp:
        speak("tell me the word sir")
        cq = takeCommand().lower()
        result=Diction.meaning(cq)
        speak(f"The Antonym of {cq} is {result}")
    speak("dictionary has been closed sir .")  
        
    
def takeCommand():
        #it takes micorphone input from the user and return strin g output 
        r=sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold=1
            #r.adjust_for_ambient_noice(source)
            #audio = r.listen(source)
            audio=r.listen(source)
            
        try:
             print("Recognizing...")
             query=r.recognize_google(audio, language='en-in')
             query=r.recognize_google(audio, language='en-in')
             print(f"Boss said:{query}\n")
        except Exception as e:
            print(e)
            #speak("say that again please...")
            print("Say that Again Please....")
            return"none"
        query=query.lower()
        return query 
    
        
def TaskExecution():
        wishMe()
        #if 1:
        while True:
            
  
            query = takeCommand().lower()
            
            #logic for ecectuing tasks based on query
            if 'wikipedia' in query:
                speak('Searching Wikipedia')
                query=query.replace("wikiepdia","")
                results=wikipedia.summary(query,sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            # elif "translator" in query:
            #     Tran()
            elif "close this tab" in query:
                keyboard.press_and_release('ctrl + w')
                speak("tab closed sir")
            elif "homepage" in query:
                keyboard.press_and_release("alt + home")
                speak("home page is here sir")
            elif "full screen mode" in query  :
                keyboard.press("f11")
                #speak("sir activated full screen mode")
            elif "close full screen mode" in query :
                keyboard.press("f11")
            elif "next video" in query:
                keyboard.press_and_release("shift + N")
                speak("moved to next video")
                #speak("deactivated full screen mode")
            elif "open new tab" in query:
                keyboard.press_and_release("ctrl + t")
                speak("opened new tab sir")
            elif "stop loading" in query:
                keyboard.press("esc")
                speak("stopped sir")
            elif "first tab" in query:
                keyboard.press_and_release("ctrl +1")
                speak("moved to first tab sir")
            elif "next tab" in query:
                keyboard.press_and_release("ctrl + 8")
                speak("moved to next tab sir")
            elif "reset browser zoom" in query or "default browser size" in query:
                keyboard.press_and_release("ctrl + 0")
                speak("browser set to its default size sir")
            elif "enter" in query:
                keyboard.press("enter")
            # elif "hide bookmark" in query : 
            #     keyboard.press_and_release("ctrl + shift + B")
            #     speak("bookmark hidden sir")
            # elif "show bookmark" in query or "unhide bookmark" in query or "visible bookmark" in query : 
            #     keyboard.press_and_release("ctrl + shift + B")
            #     speak("bookmark is now visible sir to everyone")
            elif "last tab" in query:
                keyboard.press_and_release("ctrl + 9")
                speak("switched to last tab of this browser sir")
            elif "open file in browser" in query:
                keyboard.press_and_release("ctrl + O")
                speak("opened files in browser sir")
            elif "select everything" in query or "select all" in query:
                keyboard.press_and_release("ctrl + A")
                speak("everything selected sir")
            elif "copy" in query:
                keyboard.press_and_release("ctrl + C")
                speak("copied selected item sir")
            elif "cut" in query:
                keyboard.press_and_release("ctrl + X")
                speak("cutted sir")
            elif "paste" in query:
                keyboard.press_and_release("ctrl + V")
                speak("pasted selected item sir")
            # elif "add this to bookmark" in query:
            #     keyboard.press_and_release("ctrl + D")
            #     speak("added this page to bookmark sir")
            elif "bookmark manager" in query or "book manager" in query:
                keyboard.press_and_release("ctrl + shift + O")
                speak('opened manager of this bookmark sir')
            elif "download window" in query or "show downloads" in query or "download" in query:
                keyboard.press_and_release("ctrl + J")
                speak("here is the downlaod window sir . u have downaloaded many items last few recent weeks sir .")
            elif "browser history" in query or "history" in query:
                keyboard.press_and_release("ctrl + H")
                speak("here is the browsing history sir")
                
                # while True:
                #     speak("Please tell sir do you what to clear your browsing history?")
                #     mia= takeCommand().lower()
                #     try:
                #         if "exit" in mia or "close" in mia or "no" in mia:
                #             speak("okay sir , i won't delete your browsing history sir . though i would give an suggestion sir its better to use incognito tab")
                #             break
                #         elif "yes" in mia:
                #             keyboard.press_and_release("ctrl + shift + delete")
                #             speak("ohk sir , deleting your history") #press ok
                            
                #     except:
                #         speak("not able to detect thats from whom this order is coming")
            elif "deepika" in query or "mam" in query:
                speak("Hello Deepika Mam , nice to meet u .")
            # elif "delete browsing data" in query:
            #     keyboard.press_and_release("ctrl + shift + delete")
            #     speak("ohk sir , deleting your history . just press ok .") #press ok    
            elif "new browser" in query:
                keyboard.press_and_release("ctrl + N")
                speak("opened new browser sir")
            elif "incognito mode" in query or "private mode" in query:
                keyboard.press_and_release("ctrl + shift + N")
                speak("switched to private aka incognito mode . enjoying surfing now sir")
            elif "print page" in query or "print this page" in query:
                speak("ok sir , wait printing is in progress")
                keyboard.press_and_release("ctrl + P")
                speak("page is printed sir")
            elif "refresh the page" in query or "refresh this page" in query or "refresh page" in query:
                keyboard.press("f5")
                speak("page is refreshed sir")
            elif "save the page" in query or "save this page" in query or "save the file" in query or "save this file" in query:
                keyboard.press_and_release("ctrl + S")
                speak("press ok . sir now page is saved ")
                keyboard.press("enter")
                #speak("press ok . sir now page is saved ")
            elif "enter" in query:
                keyboard.press("enter")
            elif "new tab" in query :
                keyboard.press_and_release("ctrl + T")
                speak("opened new tab sir")
            elif "source code" in query:
                keyboard.press_and_release("ctrl + U")
                speak("here is the source code of this page sir")
            elif "close this browser" in query:
                keyboard.press_and_release("ctrl + shift + W")
                speak("browser has been closed sir")
            elif "reopen" in query or "reopen tab" in query or "reopen browser" in query:
                keyboard.press_and_release("ctrl + shift + T")
                speak("we are back in the same tab sir")
            elif "close this window" in query:
                speak("ok sir closing")
                keyboard.press_and_release("alt + f4")
                speak("current window has been closed sir.")
            elif "open start" in query or "close start" in query:
                speak("ok sir")
                keyboard.press_and_release("ctrl + esc")
            elif "how are you" in query:
                speak("i am fine sir. how are u sir?")
                condition = takeCommand().lower()
                
                
                if "fine" in query or "good" in query :
                    speak("good to hear that sir")
                    
                else:
                    speak("Tomorrow’s a new day . Bad days are only temporary , the worst days can only last 24 hours . sir even the darkest night will end and sun will rise on us again sir . " )
                    pass# elif "i am fine" in query or "i am good" in query:
            #     speak("very good to hear that sir.")
            elif "boss" in query or "boss name" in query or "owner" in query:
                speak("my Boss name is RVC aka Master Rohan  Vinay  Chaudhary .")
            elif'open google' in query or "search google" in query:
                speak("sir, What should  i search on google")
                cm= takeCommand().lower()
                query = query.replace("sofia","")
                query=query.replace("google search","")
                web='https://www.google.com/search?q=' + cm
                webbrowser.open(web)
                speak("ok , sir opening google")
                ChromeAuto()
            
                
            elif "close google" in query:
                speak("okay sir, closing google")
                os.system("taskkill/f /im chrome.exe")
                speak("google closed sir")
                
            elif "ok google" in query:
                speak("got it sir , wait a second sir .")
                import wikipedia as googleScrap
                query = query.replace("sofia","")
                query = query.replace("ok google","")
                speak("This is what i found on web")
                kit.search(query)
                try:
                    result = googleScrap.summary(query,2)
                    speak(result)
                except:
                    speak("No speakable data available for this topic")
            
            #hid ethe folders
            elif "hide" in query or "hide all files" in query or"hide this folder" in query:
                os.system("attrib +h /s /d") #os module
                speak("sir ,all the files in this folder are now hidden")
            elif "unhide" in query or "visible" in query :
                os.system("attrib -h /s /d")
                speak("sir , all the files in this folder are now visible to everyone.")
             
            elif "battery status" in query or "how much power is left" in query or "how much power we have" in query:
                battery =psutil.sensors_battery()
                percentage=battery.percent
                speak(f"sir our system have {percentage} percent battery")
                if percentage >65:
                    speak("we have enough power to continue our work")
                elif 15>=percentage:
                    speak("we have very low power, please connect to charging sir , or it will shutdown the system very soon")
                elif percentage>15 and percentage <=35:
                    speak("sir i request you plug in the charger now , so we can continue our work")
            
            #elif "internet speed" in query:
              # s = speedtest.Speedtest()
               #dl=s.download()
                #up= s.upload()
                #res = s.results.dict()
                #speak(f"sir we have {dl} bit per second downlaoding speed , and {up} bit per second uploading speed")
                #try :
                 #   os.system('cmd /k "speedtest"')
                #except:
                 #   speak("internet connection is very slow")
                 #return
            elif 'power point presentation' in query:
                speak("opening Power Point presentation")
                power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
                os.startfile(power)
 
            elif "calculate" in query:
             
                app_id = "Wolframalpha api id"
                client = wolframalpha.Client(app_id)
                indx = query.lower().split().index('calculate')
                query = query.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                print("The answer is " + answer)
                speak("The answer is " + answer)
                
            elif "do some calculation" in query or "can you calculate" in query or "calculator"in query:
                r=sr.Recognizer()
                with sr.Microphone() as source:
                    speak("Sir what you want me to calcuate?")
                    print("listening...")
                    r.adjust_for_ambient_noise(source)
                    audio = r.listen(source)
                my_string=r.recognize_google(audio)
                print(my_string)
                def get_operator_fn(op):
                    return{
                        '+': operator.add, #plus
                        '-': operator.sub, #minus
                        '*': operator.mul, #multiplied by
                        'divided': operator.__truediv__,#divided
                        }[op]
                def eval_binary_expr(op1, oper, op2): #5 plus 8
                    op1,op2=int(op1), int(op2)
                    return get_operator_fn(oper)(op1, op2)
                speak("your result is")
                speak(eval_binary_expr(*(my_string.split())))
                
            elif'open website' in query:
                speak("Ok Sir , but tell me the website name sir.")
                web1 = takeCommand().lower()
                web2 = 'https://www.' + web1 + '.com'
                webbrowser.open(web2)
                speak("website Launched!")
            elif 'website' in query:
                speak("Ok Sir , Launching.....")
                query = query.replace("sofia","")
                query = query.replace("website","")
                query = query.replace(" ","")
                web1 = query.replace("open","")
                web2 = 'https://www.' + web1 + '.com'
                webbrowser.open(web2)
                speak("Launched!")  
           
            elif 'play music'in query or "play song" in query:
                Music()
                
            
            elif 'the time' in query or "time" in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")    
                speak(f"Sir, the time is {strTime}")
            elif "word file" in query or "word doc" in query :
                codePath="C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
                speak("opened word file sir")
                os.startfile(codePath)
            elif "close word file" in query:
                speak("okay sir, closing word file")
                os.system("taskkill/f /im WINWORD.EXE")
                speak("closed sir")
        
            elif 'open after effect' in query:
                speak("ok sir , opening Adobe after effects for u sir .")
                codePath="C:\\Program Files\\Adobe\\Adobe After Effects 2020\\Support Files\\AfterFX.exe"
                os.startfile(codePath)
                speak("opened sir after effects")
            elif "close after effect" in query:
                speak("okay sir, closing after effects")
                os.system("taskkill/f /im AfterFX.exe")
                speak("ok sir , closed")
                
          
            elif 'open telegram' in query:
                speak("ok sir opening telegram")
                codePath="C:\\Users\\RohanRVC\\AppData\\Roaming\\Telegram Desktop\\Telegram.exe"
                speak("ok sir , opened telegram chats")
                os.startfile(codePath)
            elif "close telegram" in query:
                speak("okay sir, closing telegram")
                os.system("taskkill/f /im Telegram.exe")
                speak("telegram closed sir")
         
            elif 'open photoshop' in query:
                speak("opening photoshop for u sir")
                codePath="C:\\Program Files\\Adobe\\Adobe Photoshop CC 2019\\Photoshop.exe"
                os.startfile(codePath)
                speak("opened photoshop sir")
            elif "close photoshop" in query:
                speak("okay sir, closing photoshop")
                os.system("taskkill/f /im Photoshop.exe")
                speak("Photoshop closed sir")
            elif 'open torrent' in query:
                codePath="C:\\Users\\RohanRVC\\AppData\\Roaming\\uTorrent\\uTorrent.exe"
                os.startfile(codePath)
                speak("torrent opened sir")
            elif "close torrent" in query:
                speak("okay sir, closing Torrent")
                os.system("taskkill/f /im uTorrent.exe")
                speak("Torrent closed sir")
                
            elif "where is" in query:
                
                Place =query.replace("where is ","")
                Place= Place.replace("sofia" , "")
                Googlemaps(Place)
                
            #to set alarm
            #to set alarm
            elif "alarm" in query:
                speak("sir tell me the time to set alarm . for example , set alarm 5:30 am ")
                tt = takeCommand()
                tt= tt.replace("set alarm ", "")
                tt=tt.replace(".","")
                tt=tt.upper()
                
                alarm(tt)
                
                #to find a joke
            elif "tell me a joke" in query or "joke" in query:
                speak("here is the joke sir")
                joke =pyjokes.get_joke()
                speak(joke)
                print(joke)
            elif "dictionary" in query:
                Dict()
            elif "read pdf" in query:
                pdf_reader()
                
            elif "shutdown the system" in query:
                speak("shutting down the system sir in less than 10 second")
                os.system("shutdown /s /t 5")
                speak(" 6 . 5 . 4 . 3 . 2 . 1 . go")
            elif "restart the system" in query:
                speak("restarting the system now")
                os.system("shutdown /r /t 5")    
                speak("see u in less then a minute sir")
            elif "hibernate now" in query or "hibernate the system" in query or "sleep the system" in query:
                speak("sleeping the system now")
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")  
            elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
                speak("Your Device has been locked sir")
            elif 'empty recycle bin' in query:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin Recycled")
            elif 'chargenull' in query:
                os.system(".")
            elif "don't listen" in query or "stop listening" in query:
                speak("for how much time you want to stop jarvis from listening commands")
                a = int(takeCommand())
                time.sleep(a)
                print(a)
            elif 'open notepad' in query:
                codePath="C:\\Users\\RohanRVC\\Documents\\notepad.txt"
                speak("opened notepad sir")
                os.startfile(codePath)
            elif "close notepad" in query:
                speak("okay sir, closing notepad")
                os.system("taskkill /f /im notepad.txt")
            elif "open command prompt" in query:
                os.system("start cmd")
                speak("here is command sir with a black screen")
            elif "close command prompt" in query:
                os.system("taskkill /f /im cmd.exe")
            #   os.startfile ("C:\\Users\\RohanRVC\\.spyder-py3\\annoying sound music.mp3")
            elif "open control panel" in query:
                os.system("start control panel ")
                speak("open all the control from here sir")
            elif "close command prompt" in query:
                os.system("taskkill /f /im control panel")
                speak("clsoed sir")
            elif "open camera" in query:
                cap=cv2.VideoCapture(0)
                while True:
                    ret, img =cap.read()
                    cv2.inshow('webcam,img')
                    k=cv2.waitKey(50)
                    if k==27:
                        break;
                        cap.release()
                    cv2.destroyAllWindows()
            elif "open mobile camera" in query:
                import urllib.request #pip install opencv-python  # pip install numpy
                import numpy as np
                URL = "http:192://192.168.43.1:8080/shot.jpg"
                while True:
                    img_arr = np.array(bytearray(urllib.request.urlopen(URL).read()),dtype=np.uint8)
                    img=cv2.imdecode(img_arr,-1)
                    q=cv2.waitKey(1)
                    if q==ord("q"):
                        break;
                cv2.destroyAllWindows()
                
            elif "weather" in query:
             
            # Google Open weather website
            # to get API of Open weather
               api_key = "Api key"
               base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
               speak(" City name ")
               print("City name : ")
               city_name = takeCommand()
               complete_url = base_url + "appid =" + api_key + "&q =" + city_name
               response = requests.get(complete_url)
               x = response.json()
               
               if x["cod"] != "404":
                   y = x["main"]
                   current_temperature = y["temp"]
                   current_pressure = y["pressure"]
                   current_humidiy = y["humidity"]
                   z = x["weather"]
                   weather_description = z[0]["description"]
                   print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                   
               else:
                       speak(" City Not Found ")
            
            # elif 'video downloader' in query:
            #     root = Tk()
            #     root.geometry('500x300')
            #     root.resizable(0,0)
            #     root.title("Youtube Video Downloader")
            #     speak("Enter Video Url Here !")
            #     Label(root,text = "Youtube Video Downloader",font = 'arial 15 bold').pack()
            #     link = StringVar()
            #     Label(root,text = "Paste Yt Video URL Here",font = 'arial 15 bold').place(x=160,y=60)
            #     Entry(root,width = 70,textvariable = link).place(x=32,y=90)
                
            #     def VideoDownloader():
            #         url = YouTube(str(link.get()))
            #         video = url.streams.first()
            #         video.download()
            #         Label(root,text = "Downloaded",font = 'arial 15').place(x= 180,y=210)
            #     Button(root,text = "Download",font = 'arial 15 bold',bg = 'pale violet red',padx = 2 , command = VideoDownloader).place(x=180,y=150)

            #     root.mainloop()
            #     speak("Video Downloaded")
                
            # elif "internet speed" in query:
            #     SpeedTest()
            # to find location
            elif "where i am" in query or "where we are" in query:
                speak("wait sir , let me check")
                try:
                    
                    ipAdd = requests.get('https://api.ipify.org').text
                    print(ipAdd)
                    url = 'https://get.geosjs.io/v1/ip/geo/'+ipAdd+'.json'
                    geo_requests = requests.get(url)
                    geo_data = geo_requests.json()
                    print(geo_data)
                    city = geo_data['city']
                    state =geo_data['state']
                    country =geo_data['country']
                    speak(f"sir we are in {city} city of {state} state of {country} country.")
                except Exception :
                     speak("sorry sir , due to networkissue i am not able to locate that were we are")
                     pass
                
            elif 'jarvis voice' in query or "jarvis" in query :
                if 'sofia' in query:
                    engine.setProperty('voice', voices[1].id)
                else:
                    engine.setProperty('voice', voices[0].id)
                    speak("Hello Sir , Jarvis here . I have switched my voice . sofia has gone to sleep sir. hope u liked  it sir.")
            
            elif "send mail" in query or "mail" in query or "email" in query :
          #      import ssl, smtplib
                speak("write down the email address sir where u wanna send mail sir")
                recipient = input("Write here sir-:")
                speak("tell the subject of the mail sir")
                subject = takeCommand()
               # speak("tell the subject of the mail sir")
              #  message=takeCommand()
                speak("tell the content u wanna write sir")
                text =takeCommand()
#smtplib is the module used by python to send emails through SMTP.
#The ssl module is used to access the Operating System Socket.

                port = 465
#This port will be required later.

#email = input("Enter your email: ")
#You key in the email address you want to send an email with.
#password = input("Enter your password: ")
#This is the password to the email inputed before

              #  recipient = input("Enter the email address to send the email to: ")
#This is the email address which the emai is being sent to.
                #subject = input("What is the subject of the email: ")
#This is the subject head of the email
               # text = input("Type your email here: ")
#This is the body of the email.
#We combine the subject and the text to form the subject head and the body.
                message = "Subject: {}\n\n{}".format(subject, text)

                context = ssl.create_default_context()
#This method just creates a new class instance for implemenmtation of a secure protocol.

                with smtplib.SMTP_SSL("smtp.gmail.com",port,context=context) as servers:
    #This syncs the local smtp server in the localhost with Gmail's server
                    servers.login("rohanrvc46@gmail.com","GudduGovindpandit@")
        #This is the login credentials being checked.
                    servers.sendmail("rohanrvc46@gmail.com", recipient,message)
                strTime = datetime.datetime.now().strftime("%H:%M:%S") 
                speak(f"mail successfully sent sir to {recipient} at {strTime}")
        #The server is initialized to send an email with the three parameters
            elif 'sofia voice' in query or "sofia" in query:
                if 'male' in query:
                    engine.setProperty('voice', voices[1].id)
                else:
                    engine.setProperty('voice', voices[1].id)
                    speak("Hello Sir ,Sofia here again at your service sir .")
            
            elif "instagram profile" in query or "open profile on instagram" in query:
                speak("sir please enter the user name correctly")
                name=input("ENter username here sir-:")
                webbrowser.open(f"www.instagram.com/{name}")
                speak(f"Sir here is the profile of the user {name}")
                
                speak("sir would you like to download profile picture of this account.")
                condition = takeCommand().lower()
                if "yes" in condition:
                    mod = instaloader.Instaloader() #pip install instadownloader
                    mod.download_profile(name,profile_pic_only=True)
                    speak("i am done sir , profile picture is saved in our main folder. now i am ready for next work")
                else:
                    pass
            
                #your
            elif "tell about yourself" in query or "sofia introduce" in query :
                speak("I am Sofia. Speed 1 terahertz, memory 1 zettabyte. A Virtual Artificial intelligiance , and i am here to assist you , with a variety of task , as best i can 24 hours a day , and 7 days a week ")
            elif "tell about yourself Jocasta" in query or "jocasta introduce" in query :
                speak("I am Jocasta. i am an backupplan of Mr stark as soon as jarvis and Friday both is disabled or he got crashed sir i will be at your service sir . Speed 1 terahertz, memory 1 zettabyte. A Virtual Artificial intelligiance , and i am here to assist you , with a variety of task , as best i can 24 hours a day , and 7 days a week ")
            
            elif "change the window" in query:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                
                pyautogui.keyUp("alt")
                speak("switched to another window sir")
            
            elif "temperature" in query:
                speak("got it sir, but which City sir?")
                cm= takeCommand().lower()
                search =("temperature in",cm) 
                url= f"https://www.google.com/search?q={search}"
                r= requests.get(url)
                data= BeautifulSoup(r.text,"html.parser")
                temp = data.find("div",class_="BNeawe").text
                speak(f"current {search} is {temp}")
                
                
            elif "take screenshot" in query or "take a screenshot" in query:
                speak("sir , what should be the name of this screenshot file?")
                name = takeCommand().lower()
                speak("please sir hold the screen for few seconds , iam taking the screenshot")
                time.sleep(3)
                img=pyautogui.screenshot()
                img.save(f"{name}.png")
                speak("i am done sir , the screenshot is been saved in our main folder . now i am ready for next work")
            
            elif "tell me news" in query or "news" in query:
                speak("please wait sir, feteching the latest news for you from all over the world sir ")
                print(speak)
                
                news()
                print(news)
            
            elif "remember that" in query or "remember mode" in query:
                speak("remember mode activated. but sir what you want me to remember")
                remeberMsg= takeCommand().lower()                
                remeber=open('data.txt','w')
                remeber.write(remeberMsg)
                remeber.close()
                speak("ok sir noted")
                
            elif "what do you remember" in query:
                remeber = open('data.txt','r')
                speak("you told me to remember that"+remeber.read())
            elif "activate how to do " in query:
                speak("How to do mode is activated ")
                while True:
                    speak("Please tell sir what to want to know?")
                    how= takeCommand()
                    try:
                        if "exit" in how or "close" in how or "deactivate" in how:
                            speak("okay sir , how to do mode is deactivated")
                            break
                        else:
                            max_results=1
                            how_to = search_wikihow(how, max_results)
                            assert len(how_to) ==1
                            how_to[0].print()
                            speak(how_to[0].summary)
                    except Exception :
                        speak("sorry sir , i am not able to find this")
                
            elif "ip address" in query:
                ip =get('https://api.ipify.org').text
                speak(f"your IP adress is{ip}")
                
            elif "send whatsapp message" in query or "text" in query:
                Whatsapp()
                
                
            elif "search youtube" in query or "search on youtube" in query:
                speak("ok sir searching on youtube .")
                speak("but sir , what should i search")
                cm= takeCommand().lower()               
                query = query.replace("sofia","")
                query=query.replace("youtube search","")
                web='https://www.youtube.com/results?search_query=' + cm
                webbrowser.open(web)
                speak("done sir")
            elif 'youtube search' in query:
                speak("OK SIR , This Is What I found For Your Search!")
                query = query.replace("sofia","")
                query = query.replace("youtube search","")
                web = 'https://www.youtube.com/results?search_query=' + query
                webbrowser.open(web)
                speak("Done Sir!")
            elif  "open youtube" in query:        
                YoutubeAuto()
            
            elif "pause" in query or "play" in query:
                keyboard.press('space bar')
                speak("done sir")
            elif "mute" in query:
                keyboard.press('m')
                speak("done sir")
            elif "skip" in query or "forward" in query:
                keyboard.press("l")
            elif "restart" in query:
                speak("restarting video sir")
                keyboard.press("0")
                speak("started playing again from start sir")
            elif "back" in query or "backward" in query:
                keyboard.press("j")
            elif "full screen" in query:
                keyboard.press("ff")
                speak("done sir")
            elif "subtitle" in query or "caption" in query:
                keyboard.press("c")
                speak("done sir")
            elif "theatre mode" in query:
                keyboard.press("t")
                speak("here is theatre mode sir enjoy")
            elif "open theatre mode" in query or "activate theatre mode"in query:
                keyboard.press("t")
                speak("here is theatre mode sir enjoy")
            elif "open theatre mode" in query or "activate theatre mode"in query:
                keyboard.press("t")
                speak("theatre mode is not deactivated sir")
            # elif "restart" in query:
            #     keyboard.press("0")
            elif "activate mini play" in query or "open mini play" in query:
                keyboard.press("i")
                speak("miniplay is here sir")
            elif "deactivate mini play" in query or "close mini play" in query:
                keyboard.press("i")
                speak("miniplay is deactivated here sir")
                
            elif "next video" in query:
                keyboard.press_and_release("shift + N")
                speak("playing next video sir")
            elif "previous" in query:
                keyboard.press_and_release("shift + P")
                speak("playing previous video")
            elif 'change background' in query:
                ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
                speak("Background changed successfully")
 
            elif 'news' in query:
             
                try:
                    jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                    data = json.load(jsonObj)
                    i = 1
                    
                    speak('here are some top news from the times of india')
                    print('''=============== TIMES OF INDIA ============'''+ '\n')
                    
                    for item in data['articles']:
                        
                        print(str(i) + '. ' + item['title'] + '\n')
                        print(item['description'] + '\n')
                        speak(str(i) + '. ' + item['title'] + '\n')
                        i += 1
                except Exception as e:
                 
                            print(str(e))
                
                    
            # Download the helper library from https://www.twilio.com/docs/python/install
    
            
            elif "send normal message" in query:
                speak("sir what message should i send")
                msz = takeCommand()
    # Find your Account SID and Auth Token at twilio.com/console
    # and set the environment variables. See http://twil.io/secure
                account_sid = os.environ['AC01f6ad1e15c6da0cd23b6e74adcb5d9d']
                auth_token = os.environ['d0d87bce38a602cf7c1c639054397554']
                client = Client(account_sid, auth_token)
    
                message = client.messages \
                    .create(
                        body=msz,
                        from_='+15017122661',
                        to='+918999425864'
                    )
    
                print(message.sid)
            
            elif "call me" in query:
                speak("sir what message should i send")
                msz = takeCommand()
    # Find your Account SID and Auth Token at twilio.com/console
    # and set the environment variables. See http://twil.io/secure
                account_sid = os.environ['AC01f6ad1e15c6da0cd23b6e74adcb5d9d']
                auth_token = os.environ['d0d87bce38a602cf7c1c639054397554']
                client = Client(account_sid, auth_token)
    
                message = client.calls \
                    .create(
                        twiml='<Response><Say>This is the second testing from sofia..</Say></Response>',
                        body=msz,
                        from_='+15017122661',
                        to='+918999425864'
                    )
    
                print(message.sid)
            
            elif "tell my location" in query:
                speak("ok , sir wait a second , sending signals to nasa !")
                speak("sir we are now connected to sattelite through NASA .")
                webbrowser.open('https://www.google.com/maps/place/VIT+BHOPAL+New+Academic+Block/@23.0750497,76.8536607,16.09z/data=!4m5!3m4!1s0x397ce9d90a8052ed:0xf16d30274bfb79f3!8m2!3d23.0785529!4d76.8500113')
                speak("sir your location is VIT BHOPAL New Academic Block , 3VH2+C2C, Unnamed Road, Kothri Kalan, Madhya Pradesh 466114 ")
            elif "repeat my words" in query:
                speak("speak sir!")
                while True:
                    cm=takeCommand()
                    if "Stop" in cm or "enough" in cm:
                        break
                    else:
                        speak(f"{cm}")
            
            elif "volume up" in query or "increase volume" in query:
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                
            elif "volume down" in query or "decrease volume" :
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                pyautogui.press("volumedown")
                
            elif "volume mute" in query:
                pyautogui.press("volumemute")
            
            elif "no thanks" in query or "u can leave now" in query or "u are free to go" in query or "take a break" in query:
                speak("okay sir, i am goin to sleep , you can call me anytime by saying wakeup")
                break 
            #speak("sir , do you have any other work?")
            
        
if __name__ == "__main__":
   while True:
        permission =takeCommand()
        if "wake up" in permission:
            TaskExecution()
        elif "quit" in permission:
            speak("thanks for using me sir , goodbye sir , have a good day")
            sys.exit()

