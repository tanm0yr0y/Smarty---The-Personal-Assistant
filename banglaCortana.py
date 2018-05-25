'''
Author:  Tanmoy Roy
Email:   roytanmoy910@gmail.com
'''
# -*- coding: utf-8 -*-
from Tkinter import *
import ttk
import webbrowser
from pygame import mixer;
import time
import pywapi
import win32com.client as wincl
import os
import string
from textblob import TextBlob
import speech_recognition as sr
import re
import ctypes
import wikipedia
import tkMessageBox
from pyavrophonetic import avro
from gtts import gTTS
import codecs
from langdetect import detect

root = Tk()
root.title('Smarty The Digital Agent')
root.geometry("400x100")
root.resizable(width = False, height = False)
root.iconbitmap('mic.ico')

style = ttk.Style()
style.theme_use('winnative')

photo = PhotoImage(file='microphone.gif').subsample(102,102)

label1 = ttk.Label(root, text='Query:')
label1.grid(row=0, column=0)
entry1 = ttk.Entry(root, width=40)
entry1.grid(row=0, column=1, columnspan=4)
btn2 = StringVar()
ok = 1

def remove_text_between_parens(data):
    ans = ''
    x = 0
    y = 0
    for i in data:
        if i == '[':
            x += 1
        elif i == '(':
            y += 1
        elif i == ']' and x > 0:
            x -= 1
        elif i == ')'and y > 0:
            y -= 1
        elif x == 0 and y == 0:
            ans += i
    return ans


def findWholeWord(w):
    return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search

def findBanglaWord(s, w):
    length = len(w)
    try:
        first_idx = s.index(w)
    except:
        return False
    last_idx = first_idx + (length - 1)
    return (first_idx == 0 or s[first_idx - 1] == ' ') and (last_idx+1 == len(s) or s[last_idx + 1] == ' ')

def callback():
    
    if btn2.get() == 'google' and entry1.get() != '':
        webbrowser.open('http://google.com/search?q='+entry1.get())
        
    elif btn2.get() == 'pipilika' and entry1.get() != '':
        webbrowser.open('https://www.pipilika.com/search?q='+entry1.get())

    elif btn2.get() == 'amz' and entry1.get() != '':
        webbrowser.open('https://amazon.com/s/?url=search-alias%3Dstripbooks&field-keywords='+entry1.get())

    elif btn2.get() == 'ytb' and entry1.get() != '':
        webbrowser.open('https://www.youtube.com/results?search_query='+entry1.get())

    else:
        pass


def get(event):

    if btn2.get() == 'google' and entry1.get() != '':
        webbrowser.open('http://google.com/search?q='+entry1.get())
        
    elif btn2.get() == 'pipilika' and entry1.get() != '':
        webbrowser.open('https://www.pipilika.com/search?q='+entry1.get())

    elif btn2.get() == 'amz' and entry1.get() != '':
        webbrowser.open('https://amazon.com/s/?url=search-alias%3Dstripbooks&field-keywords='+entry1.get())

    elif btn2.get() == 'ytb' and entry1.get() != '':
        webbrowser.open('https://www.youtube.com/results?search_query='+entry1.get())

    else:
        pass

'''
#internet connection
def isInternetOn():
    try:
        urllib2.urlopen('http://216.58.192.142', timeout=2)
        return True
    except:
        return False

def warningMsg(msg):
    tkMessageBox.showinfo("Connection Error", msg)
'''
def textToSpeech(txt):
    try:
        tts = gTTS(text = txt, lang = 'bn', slow = False)
        tts.save("result.mp3")
    except:
        print "Check your internet connection please!"

def sentence(line, querySummary):
    length = len(querySummary)
    l = 0
    summary = u""
    i = 0;
    while i < length:
        if querySummary[i] == u"।":
            l += 1
        summary += querySummary[i]
        if l == line:
            break
        i += 1
    return summary

def playSound(name):
    mixer.init()
    mixer.music.load(name)
    mixer.music.play()
    while mixer.music.get_busy():
        continue
    mixer.quit()
    
def buttonClick():
    speak = wincl.Dispatch("SAPI.SpVoice")
    recengine = "google"
    mixer.init()
    x = 0; y = 0; flag = False; mode = 0
    r = sr.Recognizer()
    global ok
    
    if ok:
        playSound("msg1.mp3")
        ok = 0
    recognised = u""
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
    playSound("start.mp3")
    f1 = False
    print u"আপনার বাক্য বলুন"
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            recognised = r.recognize_google(audio, language='bn-BD')
            f1 = True
        except sr.UnknownValueError:
            print (u"গুগল স্পিচ রিকগনিশন অডিও বুঝতে পারেনি(১)")
        except sr.RequestError as e:
            print (u"গুগল স্পিচ রিকগনিশন অডিও বুঝতে পারেনি(২)")
            #print("Could not request results from Google Speech Recognition service; {0}".format(e))


    if f1:
        print u"সফল হয়েছে"
        entry1.focus()
        entry1.delete(0, END)
        entry1.insert(0, recognised)
        if findWholeWord('hello')(recognised) or findWholeWord('hi')(recognised) or findWholeWord('hey')(recognised) or findWholeWord("what's up")(recognised):
            speak.Speak("greetings to you too. I'm ready to do your bidding.")
        elif findWholeWord('good morning')(recognised) or findWholeWord('good afternoon')(recognised) or findWholeWord('good evening')(recognised):
            speak.Speak("greetings to you too. I'm ready to do your bidding.")
        elif findBanglaWord(recognised, u"সুপ্রভাত") or findBanglaWord(recognised, u"শুভ সকাল"):
            textToSpeech(u"সুপ্রভাত! একটি চমৎকার দিন আশা করছি আপনার জন্য স্যার")
            playSound("result.mp3")
        elif findBanglaWord(recognised, u"শুভ অপরাহ্ন"):
            textToSpeech(u"ধন্যবাদ। শুভ অপরাহ্ন")
            playSound("result.mp3")
        elif findBanglaWord(recognised, u"শুভ সন্ধ্যা"):
            textToSpeech(u"ধন্যবাদ। শুভ সন্ধ্যা")
            playSound("result.mp3")
        elif findWholeWord('thank you')(recognised) or findWholeWord('thank')(recognised):
            speak.Speak("You are welcome.")
        elif findBanglaWord(recognised, u"ধন্যবাদ"):
            textToSpeech(u"আপনাকে স্বাগতম")
            playSound("result.mp3")
        elif findWholeWord('how are you')(recognised):
            speak.Speak("I am fine. Thank you")
        elif findBanglaWord(recognised, u"কেমন আছেন") or findBanglaWord(recognised, u"আছেন কেমন") or findBanglaWord(recognised, u"কেমন আছো"):
            textToSpeech(u"আমি ভালো। আপনি কেমন আছেন?")
            playSound("result.mp3")
        elif findWholeWord('your name')(recognised):
            speak.Speak("I am Smarty. I can make your work a lot easier.")
        elif findBanglaWord(recognised, u"নাম কি") or findBanglaWord(recognised, u"কি নাম"):
            textToSpeech(u"আমি Smarty")
            playSound("result.mp3")
        elif findWholeWord('open browser')(recognised) or findWholeWord('open chrome')(recognised) or findWholeWord('open firefox')(recognised) or findWholeWord('open internet explorer')(recognised):
            if findWholeWord('open browser')(recognised) or findWholeWord('browser')(recognised):
                speak.Speak("All right, I will do it for you.")
                webbrowser.open('http://www.google.com')
            elif findWholeWord('open chrome')(recognised):
                speak.Speak("All right, I will open Chrome.")
                url = 'chrome://newtab'
                chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
                webbrowser.get(chrome_path).open(url)
            elif findWholeWord('open internet explorer')(recognised) or findWholeWord('internet explorer')(recognised):
                ie = webbrowser.get(webbrowser.iexplore)
                ie.open('google.com')
            elif findWholeWord('open mozilla firefox')(recognised) or findWholeWord('open firefox')(recognised) or findWholeWord('firefox')(recognised):
                speak.Speak("All right, I will open Firefox for you.")
                webbrowser.get('C:/Program Files/Mozilla Firefox/firefox.exe %s').open('http://www.google.com')
            
        elif findWholeWord('what time is it')(recognised) or findWholeWord('what is the time now')(recognised) or findWholeWord('tell me the time')(recognised) or findWholeWord('what time')(recognised) and findWholeWord('now')(recognised):
            timenow = time.localtime()
            timeh = timenow.tm_hour; timem = timenow.tm_min
            if timeh < 13: ampm = "AM"
            else: ampm = "PM"
            if timem > 9: extra = " "
            else: extra = "oh"
            speak.Speak("the time is "+str(timeh)+extra+str(timem)+ampm)
        elif findBanglaWord(recognised, u"কয়টা বাজে") or findBanglaWord(recognised, u"বাজে কয়টা") or findBanglaWord(recognised, u"সময় কত"):
            timenow = time.localtime()
            timeH = timenow.tm_hour
            timeM = timenow.tm_min
            ampm = "AM"
            if timeH == 0:
                timeH = 12
            if timeH == 12:
                ampm = "PM"
            if timeH > 12 and timeH <= 23:
                timeH = timeH - 12
                ampm = "PM"
            
            if ampm == "AM":
                somoy = u"এখন সময় সকাল " + avro.parse(str(timeH)) + u" টা বেজে " + avro.parse(str(timeM)) + u" মিনিট"
            else:
                somoy = u"এখন সময় বিকেল " + avro.parse(str(timeH)) + u" টা বেজে " + avro.parse(str(timeM)) + u" মিনিট"
            print somoy
            textToSpeech(somoy)
            playSound("result.mp3")
        
        elif findWholeWord('what is the temperature')(recognised) or findWholeWord('what is the weather')(recognised) or findWholeWord('weather in')(recognised):
            loc = "Sylhet, Bangladesh"
            if "dhaka" in recognised:
                loc = "Dhaka, Bangladesh"
            elif "chittagong" in recognised:
                loc = "Chittagong, Bangladesh"
            elif "new york" in recognised:
                loc = "New York"
            lookup = pywapi.get_location_ids(loc)
            #print lookup
            for i in lookup:
                location_id = i
            print "loocation_id = "+location_id
            weather = pywapi.get_weather_from_weather_com(location_id)
            speak.Speak("the weather in "+loc+" is "+weather['current_conditions']['text']+". Current temp "+weather['current_conditions']['temperature']+" degrees Celsius")
        elif findBanglaWord(recognised, u"ওয়েদার কেমন") or findBanglaWord(recognised, u"আবহাওয়া কেমন") or findBanglaWord(recognised, u"আবহাওয়ার খবর") or findBanglaWord(recognised, u"তাপমাত্রা কেমন") or findBanglaWord(recognised, u"আবহাওয়ার কি অবস্থা"):
            loc = "Sylhet, Bangladesh"
            lookup = pywapi.get_location_ids(loc)
            for i in lookup:
                location_id = i

            weather = pywapi.get_weather_from_weather_com(location_id)
            text = weather['current_conditions']['text']
            temp = weather['current_conditions']['temperature']
            abohowya = u"বর্তমান আবহাওয়া "+ text + u" এবং তাপমাত্রা " + avro.parse(str(temp)) + u" ডিগ্রি সেলসিয়াস"
            textToSpeech(abohowya)
            playSound("result.mp3")
        elif findWholeWord('go to')(recognised) and ".com" in recognised:
            domain = recognised.split("go to ", 1)[1]
            speak.Speak("All right, I will take you to"+domain)
            webbrowser.open('http://www.'+domain)
        elif findWholeWord('close')(recognised):
            f = recognised.split("close ", 1)[1]
            speak.Speak(f+" is closed Sir.")
            p = 'TASKKILL /F /IM '+f+'.exe'
            try:
                os.system(p)
            except:
                speak.Speak(f+" is already closed Sir")
        elif findWholeWord('lock')(recognised) and findWholeWord('pc')(recognised) or findWholeWord('lock')(recognised) and findWholeWord('computer')(recognised) or findWholeWord('lock')(recognised) and findWholeWord('system')(recognised):
            ctypes.windll.user32.LockWorkStation()
        elif findBanglaWord(recognised, "কম্পিউটার তালা"):
            ctypes.windll.user32.LockWorkStation()
        elif findWholeWord('shutdown')(recognised) and findWholeWord('pc')(recognised) or findWholeWord('shutdown')(recognised) and findWholeWord('computer')(recognised) or findWholeWord('shutdown')(recognised) and findWholeWord('system')(recognised):     
            os.system('shutdown -s')
        elif findBanglaWord(recognised, "কম্পিউটার বন্ধ করুন") or findBanglaWord(recognised, "কম্পিউটার বন্ধ করো") or findBanglaWord(recognised, "কম্পিউটার বন্ধ কর"):
            os.system('shutdown -s')
            
        else:

            #internal search
            qDict = {}
            cntDict = {}
            for word in recognised.split():
                qDict[word] = 1
            path = u"""C:/Users/royta/Desktop/Final/"""
            files = os.listdir(path)            
            
            for txtfiles in files:
                if ".txt" in txtfiles:
                    p = path + txtfiles
                    cnt = 0
                    with codecs.open(p, 'r', encoding = 'utf-8', errors='ignore') as f:
                        for line in f:
                            for word in line.split():
                                if word in qDict:
                                    cnt+=1
                    cntDict[txtfiles] = cnt

            Max = 0
            fileName = ""
            for key, value in cntDict.iteritems():
                #print key, "-->", value
                if value > Max:
                    Max = value
                    fileName = key


            if Max == 0:
                print "Nothing matched. Looking online for solution."
            else:
                print "Maximum matched ", fileName, " --> ", Max

                
            #end of the internal search


            
            print recognised

            if detect(recognised) == "en":
                #wikipedia English stuff
                querySummary = ""
                wikipedia.set_lang("en")
                try:
                    idx = 0
                    searchResult = wikipedia.search(recognised)
                    length = len(searchResult)
                    i = 0
                    while i < length:
                        if searchResult[i] == recognised:
                            idx = i
                            break
                        i+=1

                    querySummary = wikipedia.summary(searchResult[idx], sentences = 3)
                    result = remove_text_between_parens(querySummary)
                    print result
                    speak.Speak(result)
                except IndexError:
                    playSound("error.mp3")
                except wikipedia.exceptions.DisambiguationError:
                    playSound("error.mp3")
                except wikipedia.exceptions.PageError:
                    playSound("error.mp3")
            else:
                #wikipedia Bangla stuff
                querySummary = u""
                wikipedia.set_lang("bn")
                try:
                    idx = 0
                    searchResult = wikipedia.search(recognised)
                    length = len(searchResult)
                    i = 0
                    while i < length:
                        if searchResult[i] == recognised:
                            idx = i
                            break
                        i+=1

                    playSound("waitmsg.mp3")
                    querySummary = wikipedia.summary(searchResult[idx])
                    #3 line summary
                    result = sentence(3, querySummary)
                    result = remove_text_between_parens(result)
                    #print result
                    textToSpeech(result)
                    playSound("result.mp3")
                except IndexError:
                    playSound("error.mp3")
                except wikipedia.exceptions.DisambiguationError:
                    playSound("error.mp3")
                except wikipedia.exceptions.PageError:
                    playSound("error.mp3")
    else:
        playSound("error.mp3")
    
        

entry1.bind('<Return>', get)

MyButton1 = ttk.Button(root, text='Search', width=10, command=callback)
MyButton1.grid(row=0, column=6)

MyButton2 = ttk.Radiobutton(root, text='Google', value='google', variable=btn2)
MyButton2.grid(row=1, column=1, sticky=W)

MyButton3 = ttk.Radiobutton(root, text='Pipilika', value='pipilika', variable=btn2)
MyButton3.grid(row=1, column=2, sticky=W)

MyButton4 = ttk.Radiobutton(root, text='Amz', value='amz', variable=btn2)
MyButton4.grid(row=1, column=3)

MyButton5 = ttk.Radiobutton(root, text='Ytb', value='ytb', variable=btn2)
MyButton5.grid(row=1, column=4, sticky=E)

MyButton6 = Button(root, image=photo, command=buttonClick, bd=0, activebackground='#c1bfbf', overrelief='groove', relief='sunken')
MyButton6.grid(row=0, column=5)
entry1.focus()
root.wm_attributes('-topmost', 1)
root.mainloop()


