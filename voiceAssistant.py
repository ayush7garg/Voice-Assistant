# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 21:53:20 2020

@author: Ayush Garg
"""


import speech_recognition as sr
import os
import sys
import re
import webbrowser
import smtplib
import requests
import subprocess
from pyowm import OWM
import youtube_dl
import urllib.request as urllib2
import json
from bs4 import BeautifulSoup as soup
#from urllib2 import urlopen
import wikipedia
import random
from time import strftime
import win32com.client as wincl
import pywhatkit as kit

#kit.add_driver_path("C:/Users/user/Downloads/chromedriver_win32/chromedriver.exe")

def shiroResponse(audio):
    #speaks audio passed as argument
    print(audio)
    for line in audio.splitlines():
        speak = wincl.Dispatch("SAPI.SpVoice")
        speak.Speak(audio)


def myCommand():
    #speech to text conversion
    #listens for commands
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('listening...')
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
    try:
        command = r.recognize_google(audio).lower()
        print('You said: ' + command + '\n')
    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
        print('....')
        command = myCommand();
    return command


def reddit(command):
    #open subreddit reddit
    reg_ex = re.search('open reddit (.*)', command)
    url = 'https://www.reddit.com/'
    if reg_ex:
        subreddit = reg_ex.group(1)
        url = url + 'r/' + subreddit
    webbrowser.open(url)
    shiroResponse('The Reddit content has been opened for you.')


def openWebsite(command):
    reg_ex = re.search('open (.+)', command)
    if reg_ex:
        domain = reg_ex.group(1)
        print(domain)
        url = 'https://www.' + domain + '.com'
        webbrowser.open(url)
        shiroResponse('The website you have requested has been opened for you.')
    else:
        pass


def greetings():
    day_time = int(strftime('%H'))
    if day_time < 12:
        shiroResponse('Hello. Good morning')
    elif 12 <= day_time < 17:
        shiroResponse('Hello. Good afternoon')
    else:
        shiroResponse('Hello. Good evening')


def helpMe():
    Path=os.getcwd()+'/helpMe/'
    i=1
    try:
        file_name="helpMe.txt"
        fr=open(Path+file_name,"r")
        l=fr.readlines()
        for sent in l:
            shiroResponse(sent)
        fr.close
    except IOError:
        shiroResponse("Sorry my bad. Try again")


def newsFeed():
    try:
        news_url="https://news.google.com/news/rss"
        Client=urllib2.urlopen(news_url)
        xml_page=Client.read()
        Client.close()
        soup_page=soup(xml_page,"xml")
        news_list=soup_page.findAll("item")
        for news in news_list[:15]:
            shiroResponse(news.title.text.encode('utf-8'))
    except Exception as e:
            print(e)


def time():
    import datetime
    now = datetime.datetime.now()
    shiroResponse('Current time is %d hours %d minutes' % (now.hour, now.minute))


def sendMail():
    shiroResponse('Who is the recipient?')
    recipient = myCommand()
    if 'david' in recipient:
        shiroResponse('What should I say to him?')
        content = myCommand()
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()
        mail.starttls()
        mail.login('987ayush@gmail.com', '*************')
        mail.sendmail('987ayush@gmail.com', 'amdp.hauhan@gmail.com', content)
        mail.close()
        shiroResponse('Email has been sent successfuly. You can check your inbox.')
    else:
        shiroResponse('I don\'t know what you mean!')


def playMusic():
    shiroResponse('What song shall I play?')
    mysong = myCommand()
    shiroResponse('Playing your favourite '+mysong)
    if mysong:
        kit.playonyt(mysong)


def AMA(command):
    reg_ex = re.search('tell me about (.*)', command)
    try:
        if reg_ex:
            topic = reg_ex.group(1)
            ny = wikipedia.summary(topic,sentences =3)
            shiroResponse(ny)
    except Exception as e:
            print(e)
            shiroResponse(e)


def playRhyme():
    shiroResponse('What rhyme shall I recite?')
    myRhyme = myCommand()
    shiroResponse('Playing your favourite '+myRhyme)
    if myRhyme:
        kit.playonyt(myRhyme)


def reciteCounting():
    Path=os.getcwd()+'/Countings/'
    i=1
    try:
        file_name="1.txt"
        fr=open(Path+file_name,"r")
        l=fr.readlines()
        for sent in l:
            shiroResponse(sent)
        fr.close
    except IOError:
        shiroResponse("Sorry my bad. Try again")


def assistant(command):
    #if statements for executing commands

    #open subreddit Reddit
    if 'open reddit' in command:
        reddit(command)

    #open website
    elif 'open' in command:
        openWebsite(command)

    #greetings
    elif 'hello' in command:
        greetings()

    elif 'help me' in command:
        helpMe()

    #top stories from google news
    elif 'news for today' in command:
        newsFeed()

    #time
    elif 'time' in command:
        time()

    #send email
    elif 'email' in command:
        sendMail()

    #launch yolo application
    elif 'detect object' in command:
        #appname1 = "yolo.py"
        shiroResponse('Launching object detection')
        os.startfile('I:\capstone\Voice assistant\yolo.py') 

        #shiroResponse('Launching object detection')

    #play youtube song
    elif 'song' in command:
        playMusic()

    #ask me anything
    elif 'tell me about' in command:
        AMA(command)

    elif 'rhyme' in command or 'poem' in command or 'rhymes' in command or 'poems' in command:
        playRhyme()

    elif 'counting' in command or 'countings' in command:
        reciteCounting()

    elif 'shutdown' or 'bye' in command:
        shiroResponse('Bye bye. Have a nice day')
        sys.exit()

    else:
        shiroResponse("Sorry. My bad. Didn't get that. Can you please repeat?")


shiroResponse('Hi, I am Shiro and I am your friend and teacher, Please give a command or say "help me" and I will tell you what all I can do for you.')

#loop to continue executing multiple commands
while True:
    assistant(myCommand())
