'''Created on Sun Aug 16 17:37:52 2020
@author: Ajay'''

import pytesseract 
import cv2 as cv
import numpy as np
from PIL import Image,ImageEnhance

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import speech_recognition as sr
import pyaudio
from gtts import gTTS
import pyttsx3

from playsound import playsound

import os
import time
import random

import requests

import datetime
import calendar

import wmi

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

import re

rec=sr.Recognizer()
options=Options()
options.binary_location='C:\\Program Files (x86)\\BraveSoftware\\Brave-Browser\\Application\\brave.exe'
months={'1':'January','2':'February','3':'March','4':'April','5':'May','6':'June','7':'July','8':'August','9':'September','10':'October','11':'November','12':'December'}

# def t_s(text):
#     spee=gTTS(text,lang='en')
#     spee.save('a_file.mp3')
#     playsound('a_file.mp3')
#     os.remove('a_file.mp3')
def ts(text):
    engine = pyttsx3.init()
    engine.setProperty('rate', 140)     # setting up new voice rate
    engine.setProperty('volume',1.0)
    engine.say(text)
    engine.runAndWait()

var=True
while var:
    with sr.Microphone() as so:
        while var:
            print('say my name...')

            try:
                rec.adjust_for_ambient_noise(so)
                name=rec.listen(so)
                name=rec.recognize_google(name,language='en-IN')
                name2=name.upper()
                name_list=['CROMA','ROOMA','AROMA','ROMA','RAMA']
                if any(s_s in name2 for s_s in name_list):

                    print(name2)
                    y=["ya, i'm listening",'hey sir!','Hey! what can i do for you',"what's your command sir",'ya, what are your orders','hello sir','what can i do for you?',"yes sir!"]
                    answer=random.choice(y)
                    # tts=gTTS(answer,lang='en')
                    playsound('tones_for_roma\\ting(1).mp3')
                    # tts.save('s.mp3')
                    # playsound('s.mp3')
                    
                    # os.remove('s.mp3')
                    ts(answer)
                    print(answer)
                    while True:
                        try:
                            print('order me sir!')
                            # with sr.Microphone() as so:
                            rec.adjust_for_ambient_noise(so,0.6)
                            command=rec.listen(so,10)
                            command=rec.recognize_google(command,language='en-IN')
                            name2=command.upper()
                            name3=name2.split(' ')
                            print(name3)
                            date=datetime.date

                            if 'PLAY' in name3:
                                n2=name2.lstrip('PLAYING')
                                print(n2)
                                browser=webdriver.Chrome(options=options,executable_path='F:\\chromedriver.exe')                            
                                browser.get('https://music.youtube.com/')
                                cook=WebDriverWait(browser,5).until(ec.element_to_be_clickable((By.XPATH,"/html/body/ytmusic-app/ytmusic-app-layout/ytmusic-nav-bar/div[2]/ytmusic-search-box/div/div[1]/paper-icon-button[1]")))
                                cook.click()
                                find=WebDriverWait(browser,5).until(ec.element_to_be_clickable((By.XPATH,"/html/body/ytmusic-app/ytmusic-app-layout/ytmusic-nav-bar/div[2]/ytmusic-search-box/div/div[1]/input")))
                                find.send_keys(n2)
                                find.send_keys(Keys.ENTER)
                                songg=WebDriverWait(browser,5).until(ec.element_to_be_clickable((By.XPATH,"/html/body/ytmusic-app/ytmusic-app-layout/div[3]/ytmusic-search-page/ytmusic-section-list-renderer/div[2]/ytmusic-shelf-renderer[1]/div[2]/ytmusic-responsive-list-item-renderer/div[1]/ytmusic-item-thumbnail-overlay-renderer/div")))
                                songg.click()
                            elif 'TELL' in name3 and 'ABOUT' in name3 or 'SHOW' in name3 and 'ABOUT' in name3:
                                wiki_object=name3[-1].lower().capitalize()
                                print(wiki_object)
                                browser=webdriver.Chrome(options=options,executable_path='F:\\chromedriver.exe')                            
                                hh=browser.get("https://en.wikipedia.org/wiki/{}".format(wiki_object))
                                hh=browser.find_element_by_xpath('//*[@id="mw-content-text"]/div[1]/p[3]').text
                                ii=re.findall(r'\[.*?\]', hh) 
                                for i in range(len(ii)):
                                    hh=hh.replace(ii[i],'')
                                ts(hh)
                            elif 'WAIT' in name3 or 'RESUME' in name3 or 'PAUSE' in name3 or 'CONTINUE' in name3:
                                pause=browser.find_element_by_xpath('//*[@id="player"]/div[4]')
                                pause.click()
                            elif 'TOP' in name3 or 'CANCEL' in name3 or 'STOP' in name3 or 'CLOSE' in name3:
                                browser.close()
                                print('Closed')
                            elif 'WEATHER' in name3 or 'WHETHER' in name3: #may use NER like spacy to extract out city names.
                                weth=name3
                                if weth[-1]=='WEATHER':
                                    city=weth[-2]
                                    city=city.replace("'S","")
                                    print(city)
                                else:
                                    city=weth[-1]
                                    print(city)
                                ress=requests.get("https://api.openweathermap.org/data/2.5/weather?q={}&appid=19fa331af660358bbb5b37c775c90e84&units=metric".format(city))
                                ress=ress.json()
                                # print(ress)
                                ress1,ress2=ress['main'],ress['weather']
                                ress2=ress2[0]
                                # yy2=yy['weather']
                                temperature,feels_like,humidity,condition=ress1['temp'], ress1['feels_like'],ress1['humidity'],ress2['main']
                                text=""" Right now in  {}  its  {} , the temperature is {} degree celsius, 
                                But it feels like {} degrees because of current humidity of {} percent""".format(city.lower().capitalize(),condition,temperature,feels_like,humidity)
                                print(text)
                                ts(text)
                            elif 'DATE' in name3 and 'IS' in name3 or 'AAJ' in name3 and 'TARIKH' in name3:
                                my_new_date= date.today().strftime("%d %m %Y")
                                my_new_date= my_new_date.split()
                                month,date,year=my_new_date[1].lstrip('0'),my_new_date[0].lstrip('0'),my_new_date[2]
                                month=months[month]
                                full_date='Today is {} {} {}'.format(month,date,year)
                                ts(full_date)
    
                            elif 'DAY' in name3 and 'IS' in name3:
                                # print(name3)
                                # date=datetime.date
                                my_date= date.today()
                                weekday=calendar.day_name[my_date.weekday()]
                                ts(weekday)
                            elif 'KAUN' in name3 or 'KYA' in name3 and 'DIN' in name3:
                                # date=datetime.date
                                # print(name3)
                                my_date= date.today()
                                weekday=calendar.day_name[my_date.weekday()]
                                ts(weekday)
                            elif 'SCREEN' in name3 and 'BRIGHTNESS' in name3 or 'BRIGHTNESS' in name3 and 'LEVEL' in name3 or 'BRIGHTNESS' in name3 and 'TO' in name3:
                                for word in name3:
                                    if word.endswith('%'):
                                        word=int(word.rstrip('%'))
                                        c = wmi.WMI(namespace='wmi')
                                        methods = c.WmiMonitorBrightnessMethods()[0]
                                        methods.WmiSetBrightness(word, 0)
                            elif 'SOUND' in name3 and 'LEVEL' in name3 or 'VOLUME' in name3 and 'LEVEL' in name3 or 'SOUND' in name3 and 'TO' in name3:
                                for word in name3:
                                    if word.isdigit():
                                        word=int(word)
                                        devices = AudioUtilities.GetSpeakers()
                                        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                                        volume = cast(interface, POINTER(IAudioEndpointVolume))
                                        volume.GetMute()
                                        volume.GetMasterVolumeLevel()
                                        volume.GetVolumeRange()                        
                                        volume.SetMasterVolumeLevel(-word, None)
                            elif 'READ' in name3 and 'THIS' in name3 or 'DOCUMENT' in name3 and 'READ' in name3 or 'READ' in name3 and 'PAGE' in name3:
                                # print('reading mode ON!....')
                                ts('reading mode has started Sir')
                                # playsound('tones_for_roma\\reading.mp3')
                                # t=cv.VideoCapture("https://192.168.0.117:8080/video")
                                # while 1:
                                #     _,ff=t.read()
                                #     cv.namedWindow('bouy',cv.WINDOW_NORMAL)
                                #     cv.imshow('bouy',ff)
                                #     k=cv.waitKey(5)
                                #     if k==27:
                                #         cv.imwrite('doc.jpg',ff)
                                
                                #         break 
                                # cv.destroyAllWindows()    
                                # t.release()
                                
                                image= cv.imread('C:\\Users\\Ajay\\Downloads\\IMG_20200913_154829.jpg')
                                x,y,z=image.shape
                                i=image
                                
                                i=cv.cvtColor(i,cv.COLOR_BGR2GRAY)                                
                                r,c=i.shape
                                print(r,c)
                                ratio_r,ratio_c=round(r/5.4),round(c/4.8)
                                c1,c2=round(r/2),round(c/2)
                                img=i[c1-int(ratio_r/2):c1+int(ratio_r/2),c2-int(ratio_c/2):c2+int(ratio_c/2)]
                                img2=img.reshape(-1)
                                
                                img3=np.sort(img2,axis=0)[:50]
                                
                                thr=img3.mean()
                                # thr=min(img2)
                                thr=thr.item()
                                thr=thr+50
                                print(thr)
                                # print(thrs)
                                # print(type(thrs))
                                
                                _,im=cv.threshold(i,thr,255,cv.THRESH_TOZERO)
                                # print(type(im))
                                # cv.imshow('fdfgd',im)
                                # cv.waitKey(0)
                                # cv.destroyAllWindows()
                                boxes=pytesseract.image_to_data(im)
                                # print(boxes)
                                row_s=[]
                                col_s=[]
                                for m,b in enumerate(boxes.splitlines()):
                                    if m!=0:
                                       b=b.split()
                                       if len(b)==12:
                                           x,y,w,h=int(b[6]),int(b[7]),int(b[8]),int(b[9])
                                           if x<10 or y<10 or w>c-10 or h>x-10:
                                               pass
                                           else:
                                           # print(x,y,w,h)
                                               row_s.append(y)
                                               col_s.append(x)
                                               cv.rectangle(im,(x-1,y-1),(x+w+20,y+h+2),(0,0,255),5)
                                min_r=min(row_s)
                                max_r=max(row_s)
                                min_c=min(col_s)
                                max_c=max(col_s)
                                cv.rectangle(i,(min_c-10,min_r-10),(max_c+35,max_r+30),(0,255,0),3)
                                i_cropped=i[min_r-10:max_r+30,min_c-10:max_c+30]
                                # _,i_cropped=cv.threshold(i_cropped,thr-20,255,cv.THRESH_TOZERO)
                                print(type(i_cropped))
                                pil_im=Image.fromarray(i_cropped)
                                enha=ImageEnhance.Contrast(pil_im)
                                factor=1.5
                                contrasted=enha.enhance(factor)
                                np_img=np.array(contrasted)
                                text=pytesseract.image_to_string(np_img)
                                print(text)
                                # speak=gTTS(text,lang='en')
                                # speak.save('book_reading.mp3')
                                # playsound('book_reading.mp3')
                                # os.remove('book_reading.mp3')
                                ts(text)
                                row_s.sort()
                                col_s.sort()
                                print(row_s)
                                print(col_s)
                                cv.namedWindow('ff',cv.WINDOW_NORMAL)
                                cv.namedWindow('gg',cv.WINDOW_NORMAL)
                                # cv.resizeWindow()
                                cv.imshow('ff',image)
                                cv.imshow('gg',i_cropped)
                                
                                cv.waitKey(0)
                                cv.destroyAllWindows()

                            elif 'DOWNLOAD' in name3 and 'MOVIE' in name3:
                                   try:
                                       ts('which movie sir?') 
                                       # playsound('tones_for_roma\\movie.mp3')
                                       options.add_experimental_option('debuggerAddress','localhost:5655')
                                       rec.adjust_for_ambient_noise(so,0.5)
                                       movie_name=rec.listen(so)
                                       movie_name=rec.recognize_google(movie_name,language='en-IN')
                                       movie_name=movie_name.split(' ')
                                       movie_name='+'.join(movie_name)
                                       print(movie_name)
                                       # brave.exe -remote-debugging-port=5655 --user-data-dir="F:\selenium_testing\brave_testing"
                                       brow=webdriver.Chrome(options=options,executable_path='F:\\chromedriver.exe')    
                                       brow.get("https://torrentzeu.org/data.php?q={}".format(movie_name))
                                       seeds=brow.find_element_by_xpath('//*[@id="table"]/tbody/tr[1]/td[2]')
                                       seeds=str(seeds.text)
                                       size=brow.find_element_by_xpath('//*[@id="table"]/tbody/tr[1]/td[3]')
                                       size=str(size.text)
                                       text="The size of the movie is {}, and it has {} seeds. Downloading will start now!".format(size,seeds)
                                       ts(text)
                                       # rec.adjust_for_ambient_noise(so)
                                       # seed_and_size=rec.listen(so)
                                       # print('im listening')
                                       # seed_and_size=rec.recognize_google(seed_and_size,language='en-IN')
                                       # print(seed_and_size)
                                       # seed_and_size=seed_and_size.upper()
                                       # seed_and_size=name2.split(' ')
                                       # if 'YES' in seed_and_size or 'HAAN' in seed_and_size or 'OK' in seed_and_size or 'HAN' in seed_and_size:
                                       # print(seed_and_size)
                                       nu=brow.find_element_by_xpath('//*[@id="table"]/tbody/tr[1]/td[5]')
                                       nu.click()
                                       # elif 'NO' in seed_and_size or 'NA' in seed_and_size or 'NAHI' in seed_and_size:
                                       # print('ok sir!')
                                           
                                       # time.sleep(1)
                                       # nu.click()
                                   except Exception as er:
                                             print(er)
                                             continue
                            elif 'CHAHAT' in name3 or 'CHACHA' in name3 or 'CHAND' in name3 and 'BIRTHDAY' in name3:
                                  playsound('tones_for_roma\\chahat_bday.mp3')
                            elif 'JYOTI' in name3 and 'BIRTHDAY' in name3:
                                  playsound('tones_for_roma\\jyoti_bday.mp3')
                            elif 'AJAY' in name3 and 'BIRTHDAY' in name3:
                                  playsound('tones_for_roma\\my_bday.mp3')
                            elif 'TUSHAR' in name3 and 'BIRTHDAY' in name3:
                                  ts(text='August 30th')
                            elif 'SHASHANK' in name3 and 'BIRTHDAY' in name3:
                                  ts(text='July 24th')
                            elif 'NISHU' in name3 and 'BIRTHDAY' in name3:
                                  ts(text='June 4th')
                            elif 'KRISHU' in name3 and 'BIRTHDAY' in name3:
                                  ts(text='February 2nd')
                            elif 'NEHA' in name3 and 'BIRTHDAY' in name3:
                                  ts(text='March 9th')
                            elif 'GO' in name3 or 'TERMINATE' in name3 or 'SLEEP' in name3 or 'BYE' in name3 or 'BYE-BYE' in name3 or 'BHAG' in name3:
                                  text='Bye Bye Sir! See you again!'
                                  ts(text)
                                  # playsound('tones_for_roma\\bye.mp3')
                                  var=False
                                  break
                                           
                                   
                            else:
                                continue
                        except Exception as e:
                            print(e)
                            continue
                else:
                    print(name2, 'again....')
            except Exception as e:
                print(e)
                continue