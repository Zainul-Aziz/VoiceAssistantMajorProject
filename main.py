#Importing all necessary libraries

from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pytz
import tkinter as tk
from tkinter import ttk
import webbrowser
import pywhatkit
from tkinter import *
import speech_recognition as sr
import pyttsx3 as p
import wikipedia
import smtplib
from email.message import EmailMessage
import sys
import time
import json
import requests
import pyautogui
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import win32com.client as wincl
from urllib.request import urlopen
from newsapi.newsapi_client import NewsApiClient
import pyjokes
import datetime
from ecapture import ecapture as ec
import os
import sports
from PIL import ImageTk, Image
from resources import MyAlarm

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january", "february", "march", "april", "may", "june","july", "august", "september","october","november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENTIONS = ["rd", "th", "st", "nd"]

root = tk.Tk()
root.title("Jarvis 0.0.1 ")

root.geometry("358x129")
# root.iconbitmap('C:\\Users\\lenovo\\PycharmProjects\\MajorProject\\myicon.ico')
root.iconbitmap("resources/myicon.ico")
canvas=Canvas(root, width=300, height=125)
img=ImageTk.PhotoImage(Image.open("resources/JARVISs.png"))

canvas.create_image(0, 0, anchor=NW, image=img)
canvas.pack()
# Initialise tkinter engine
engine = p.init()
engine.setProperty("rate", 180)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

# Weather api id
app_id = "API_KEY"


def talk(text):
    # print(f'Alexa: {text} ')
    engine.say(text)
    engine.runAndWait()

style = ttk.Style()
style.theme_use('winnative')

query = ttk.Label(root, font=("Arial bold", 9), text='Query : ')
query.place(x=27, y=98)

entry1 = ttk.Entry(root,width=30)
entry1.place(x=75, y=98)

###--------                Whatsapp Web Automation Implementation            ------------###

#Function for getting user from search box(i.e. if not in recent chat)
def new_chat(user_name,browser):
    new_chat = browser.find_element_by_xpath('//div[@class="_3LX7r"]')
    new_chat.click()

    # Enter the name of chat
    new_user = browser.find_element_by_xpath('//div[@class="_2_1wd copyable-text selectable-text"]')
    new_user.send_keys(user_name)

    time.sleep(1)

    try:
        # Select for the title having user name
        user = browser.find_element_by_xpath('//span[@title="{}"]'.format(user_name))
        user.click()
    except NoSuchElementException:
        print('Given user "{}" not found in the contact list'.format(user_name))
    except Exception as e:
        # Close the browser
        browser.close()
        print(e)
        sys.exit()


def whatsapp_chat(user):
    chrome_browser=webdriver.Chrome()
    chrome_browser.get('https://web.whatsapp.com/')
    time.sleep(15)
    user_name_list = [user]

    for user_name in user_name_list:

        try:
            # Select for the title having user name
            user = chrome_browser.find_element_by_xpath('//span[@title="{}"]'.format(user_name))
            user.click()
        except NoSuchElementException as se:
            new_chat(user_name,chrome_browser)
            pass

        # Typing message into message box
        message_box = chrome_browser.find_element_by_xpath('//div[@class="_2A8P4"]')
        talk("Please tell What is your message")
        print("Alexa: Please tell What is your message")
        message = takeCommand()
        time.sleep(5)
        message_box.send_keys(message)

        # Click on send button
        message_box = chrome_browser.find_element_by_xpath('//button[@class="_1E0Oz"]')
        message_box.click()

    chrome_browser.close()
    pass


###--------        Email Sending Implementation         ------------###

email_list = {
    'Ajmal Khan': 'EMAILID@gmail.com',
    'I': 'EMAILID@gmail.com',
    'Hammad': 'EMAILID@gmail.com',
    'Sayeed': 'EMAILID@gmail.com',
    'Faiz': 'EMAILID@gmail.com'
}


def send_email(receiver, subject, message):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    # Make sure to give app access in your Google account
    server.login('EMAIL_ID', 'PASSWORD')
    email = EmailMessage()
    email['From'] = 'Sender_Email'
    email['To'] = receiver
    email['Subject'] = subject
    email.set_content(message)
    try:
        server.send_message(email)
        return "Success"
    except:
        return "Failed"


def get_email_info():
    talk('To Whom you want to send email')
    print('Alexa: To Whom you want to send email')
    name = takeCommand()
    receiver = email_list[name]
    print(receiver)
    talk('Tell me the subject of your email?')
    print('Alexa: Tell me the subject of your email?')
    subject = takeCommand()
    talk('Tell me the text of your email?')
    print('Alexa: Tell me the text of your email?')
    message = takeCommand()
    if send_email(receiver, subject, message)=='Success':
        talk('Hello sir, Your email is sent')
        print('Alexa: Hello sir, Your email is sent')
    else:
        talk('Something went wrong Please try again')
        print('Alexa: Something went wrong Please try again')


###--------        Weather Forecast Implementation         ------------###


def weather_and_temperature(query):
    words = query.split()
    city = str(words[-1])
    complete_api_link = "http://api.openweathermap.org/data/2.5/weather?q=" + city + "&appid=" + app_id + ""
    api_link = requests.get(complete_api_link)
    api_data = api_link.json()
    if api_data['cod'] == '404':
        print("Invalid City: {},Please check your city name ".format(city))
        talk("Alexa: Invalid City: {},Please check your city name ")
    else:
        temp_city = ((api_data['main']['temp']) - 273.15)
        weather_desc = api_data['weather'][0]['description']
        #    humidity = api_data['main']['humidity']
        #    wind_speed = api_data['wind']['speed']
        talk("Its currently " + weather_desc + " and {:.0f} degrees".format(temp_city) + " in " + city)
        print("Alexa: Its currently " + weather_desc + " and {:.0f} ℃".format(temp_city) + " in " + city)


###--------  Google Calendar Automation   ----###


def google_authentication():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('resources/token.json'):
        creds = Credentials.from_authorized_user_file('resources/token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'resources/client_secret.json', SCOPES)
                # Create google developer console project and obtain client secret json file
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('resources/token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)
    return service


def get_events(day, service):
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.timezone("Asia/Calcutta")
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        talk('No upcoming events found.')
        print('Alexa: No upcoming events found.')
    else:
        talk(f"You have {len(events)} events on this day.")
        print(f"Alexa: You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))

            start_time = str(start.split("T")[1].split("-")[0])
            if int(start_time.split(":")[0]) < 12:
                # start_time = start_time.replace("+05:30"," ")
                # start_time = start_time.replace(":00"," ") + "am"
                start_time= start_time.split(":")[0] + ":" + start_time.split(":")[1] + " am"
            else:
                start_time = str(int(start_time.split(":")[0])-12) + ":" + start_time.split(":")[1] + " pm"
                # start_time = start_time.replace("+05:30"," ") + "pm"
                # start_time=start_time + "pm"
            eventDetails = event["summary"] + " at " + start_time
            print("Alexa: " + eventDetails)
            talk(event["summary"] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENTIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    if month < today.month and month != -1:
        year = year+1


    if month == -1 and day != -1:
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month


    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)

    if day != -1:  # FIXED FROM VIDEO
        return datetime.date(month=month, day=day, year=year)



def searchLocation(loc):
    chrome_browser=webdriver.Chrome()
    chrome_browser.get("Put Google Map Site Link here")
    Place=chrome_browser.find_element_by_class_name("tactile-searchbox-input")
    Place.send_keys(loc)
    Submit=chrome_browser.find_element_by_xpath('//*[@id="searchbox-searchbutton"]')
    Submit.click()



###--------        Search query in input field of application         ------------###


def callback():
    webbrowser.open('http://google.com/search?q='+entry1.get())


def get(event):
    webbrowser.open('http://google.com/search?q='+entry1.get())

greeting=1
def takeCommand(): #taking audio and returning string


    r=sr.Recognizer()

    with sr.Microphone () as source:
        print(f"Listening")
        r.pause_threshold = 0.6
        r.adjust_for_ambient_noise(source,duration=1)
        text = r.listen(source)
        recognisedText=""
        try:
            recognisedText= r.recognize_google(text)
            print(f"Recognizing...")
            print(f"User: {recognisedText}")
            return recognisedText
        # except:
        #     pass
        #     talk("Sorry didnot understand that")
        except sr.UnknownValueError:
            print("")
            return "try again"
        # except sr.RequestError as e:
        except sr.RequestError:
            return "try again"
            # print("Alexa: Say that again please")

counter=0


#username = takeCommand()

if True:

    def buttonClick():

        newsapi = NewsApiClient(api_key="NEWS_API_KEY")
        # talk("Any command")
        # def speak(audio):
        #     engine.say(audio)
        #     engine.runAndWait()

        def wishMe():
            hour = int(datetime.datetime.now().hour)
            # talk("")
            if hour>=0 and hour<12:
                talk(" Good Morning Sir")

            elif hour>=12 and hour<17:
                talk(" Good Afternoon Sir")

            else:
                talk(" Good Evening Sir")
            talk("Any command")
            print("Alexa: Any command")



        if __name__ == "__main__":
            global recognised_text
            wishMe()
            if 1:

                try:
                    recognised_text = takeCommand().lower()
                    #Search wilipedia
                    if 'wikipedia' in recognised_text:
                        talk("searching on wikipedia sir")
                        print("Alexa: searching on wikipedia sir")
                        recognised_text = recognised_text.replace("wikipedia", "")
                        results = wikipedia.summary(recognised_text, sentences=2)
                        print(f'Alexa: {results}')
                        talk(results)

                    #Open Youtube
                    elif 'open youtube' in recognised_text :
                        webbrowser.open("https://www.youtube.com/")

                    #Open Google
                    elif 'open google' in recognised_text:
                        webbrowser.open("https://www.google.com")

                    elif 'google search' in recognised_text or 'search google' in recognised_text:
                        talk("What do you want to search")
                        searchItem=takeCommand()
                        webbrowser.open("https://www.google.com/search?q="+searchItem)

                #Open Firefox
                    elif 'open firefox' in recognised_text:
                        codePath = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"
                        os.startfile(codePath)

                #Open Music on youtube
                # elif 'play music' in recognised_text:
                #     webbrowser.open("https://www.youtube.com/watch?v=36YnV9STBqc")

                #Search an item on amazon
                    elif "amazon" in recognised_text or "online shopping" in recognised_text:
                        talk("what do you want to search sir")
                        print("Alexa: what do you want to search sir")
                        item = takeCommand()
                        webbrowser.open('https://amazon.com/s/?url=search-alias%3Dstripbooks&field-keywords='+item)

                #Open Desktop Notepad
                    elif 'open notepad' in recognised_text:
                        os.system("Notepad")

                #Tell the time and date
                    elif 'date and time' in recognised_text or 'time and date' in recognised_text or 'date time' in recognised_text:
                        strTime = datetime.datetime.now().strftime("%H:%M:%S")
                        date = datetime.datetime.now().date().strftime("%A, %d %b %Y")
                        print(f"Alexa: Time is {strTime}, Date is {date}")
                        talk(f"sir the time is {strTime} and date is {date}")

                    elif 'time' in recognised_text:
                        strTime = datetime.datetime.now().strftime("%H:%M:%S")
                        talk(f"sir the time is {strTime}")
                        print(f'Alexa: sir the time is {strTime}')

                    elif 'date' in recognised_text:
                        date = datetime.datetime.now().date().strftime("%A, %d %b %Y")
                        print(f"Alexa: {date}")
                        talk('Current date is' + date)

                #Capture photo
                    elif "photo" in recognised_text or "take a photo" in recognised_text:
                        date = datetime.datetime.now()
                        imgname=(str(date).replace(":","_"))[:19]
                        ec.capture(0, "Jarvis Camera ", "resources/img_" + imgname + ".jpg")
                        talk("Task completed")
                        print("Alexa: Task Completed")

                #Open online calculator on Google
                    elif 'calculator' in recognised_text or 'calculation' in recognised_text:
                        webbrowser.open("https://www.google.com/search?q=calculator")

                #Save a note locally on your system
                    elif "note" in recognised_text:
                        talk("What should i write, sir")
                        print("Alexa: What should i write, sir")
                        note = takeCommand()
                        date = datetime.datetime.now()
                        filename="resources/note_"+(str(date).replace(":", "_"))[:19] + ".txt"

                        file = open(filename, 'w')
                        talk("Sir, Should i include date and time")
                        print("Alexa: Sir, Should i include date and time")
                        snfm = takeCommand()
                        if 'yes' in snfm or 'sure' in snfm:
                            strDate = datetime.datetime.now().date().strftime("%A, %d %b %Y")
                            strTime = datetime.datetime.now().strftime("%H:%M:%S")
                            file.write(strDate + " " + strTime)
                            file.write(" :- ")
                            file.write(note)
                        else:
                            file.write(note)
                        talk("Task completed.")
                        print("Alexa: Task Completed")
                    #Tell top news headlines
                    elif 'news' in recognised_text or 'headlines' in recognised_text:
                        try:
                            jsonObj = urlopen('https://newsapi.org/v2/top-headlines?'
                                    'country=in&'
                                    'apiKey=API-KEY')
                            data = json.load(jsonObj)
                            i = 1

                            talk('here are some top news from india')
                            print('Alexa: Here are some top news from India')
                            # print('''=============== TIMES OF INDIA ============'''+ '\n')  #returns 20 top news

                            for item in data['articles']:
                                print(str(i) + '. ' + item['title'] + '\n')
                                print(item['description'] + '\n')
                                talk(str(i) + '. ' + item['title'] + '\n')
                                i += 1
                                if i == 6:
                                    break
                        except  Exception as e:
                            print(str(e))

                    #My Code
                    #Play music on youtube
                    elif 'play' in recognised_text:
                        song = recognised_text.replace('play', '')
                        print('Alexa: Playing '+song)
                        talk('playing ' + song)
                        pywhatkit.playonyt(song)

                    #Search for person's wikipedia page
                    elif 'search for' in recognised_text:
                        person = recognised_text.replace('search for', '')
                        info = wikipedia.summary(person, 2)
                        print(f'Alexa: {info}')
                        talk(info)

                    #Tell jokes
                    elif 'joke' in recognised_text:
                        joke = pyjokes.get_joke()
                        print(f'Alexa: {joke}')
                        talk(joke)

                    #Speak and display weather forecast
                    elif 'weather' in recognised_text or 'temperature' in recognised_text:
                        weather_and_temperature(recognised_text)

                    #Whatsapp automation (send message)
                    elif 'whatsapp' in recognised_text or 'message' in recognised_text:
                        print('Alexa: With whom you want to chat')
                        talk("With whom you want to chat")
                        username = takeCommand()
                        whatsapp_chat(username.capitalize())
                        print("Alexa: Message sent successfully")
                        talk("Message sent successfully")

                #Send an email to recognised persons
                    elif 'mail' in recognised_text or 'email' in recognised_text:
                        get_email_info()

                    #Take screesshot
                    elif 'screenshot' in recognised_text:
                        screenshot=pyautogui.screenshot()
                        # file_location=asksaveasfilename()
                        date = datetime.datetime.now()
                        filename=(str(date).replace(":","_"))[:19]
                        screenshot.save('resources/screenshot_' + filename +'.png')
                        talk("Task completed")
                        print("Alexa: Task Completed")
                    # Calendar Automation
                    elif "what do i have" in recognised_text or "am i busy" in recognised_text or "do i have plans" in recognised_text or "calendar" in recognised_text:
                        SERVICE = google_authentication()
                        date = get_date(recognised_text)
                        if date:
                            get_events(date, SERVICE)
                        else:
                            talk("Please Try Again.")
                    # Set Alarm
                    elif 'alarm' in recognised_text:
                        talk("Please tell time to set an alarm ?")
                        print("Alexa: Please tell time to set an alarm ?")
                        queries = takeCommand()
                        time = queries.replace(".", "")
                        time = time.upper()
                        print(time)
                        MyAlarm.alarm(time)

                    elif "search location" in recognised_text or "find location" in recognised_text:
                        # query = recognised_text.replace("where is", "")
                        # location = query
                        talk("Which location")
                        location = takeCommand().lower()
                        searchLocation(location)





                #other questions
                    elif 'open gogoanime' in recognised_text:
                        webbrowser.open("gogoanime.com")

                    elif 'portfolio' in recognised_text:
                        webbrowser.open("https://zainul-aziz.github.io/Portfolio/index.html")

                    elif 'favourite superstar' in recognised_text:
                        talk("Alexa: Shahrukh Khan is the greatest superstar of all time")

                    elif "i love you" in recognised_text :
                        talk("Alexa: Love is hard to understand")

                    elif "how are you" in recognised_text :
                        talk("Alexa: I'm fine sir,Alhamdulillah")

                    elif "who made you" in recognised_text :
                        talk("Alexa: I am Alexa i am created by Legend")

                    elif "who are you" in recognised_text :
                        talk("Alexa: I am your virtual assistant,Alexa")

                    elif 'thank you' in recognised_text:
                        talk("Alexa: Your most welcome")

                    elif "who i am" in recognised_text:
                        talk("Alexa: If you talk then definitely your human.")

                    elif 'quit' in recognised_text or 'over' in recognised_text:
                        talk("Alexa: Ok bye Have a nice day")
                        return 0
                    # elif 'try again' in recognised_text:
                    #     talk("Sorry i did not understand that Please try again.")
                    #     buttonClick()
                    elif 'score' in recognised_text:
                        matches = sports.get_sport(sports.CRICKET)

                        #printing all matches
                        for item in matches:
                            print(item)
                    # else:
                    #     talk("Sorry i did not understand")
                    #     # talk("Sorry i did not understand that Please try again.")
                    #     # buttonClick()
                except:
                    talk("Sorry i did not understand that Please try again.")
                    print("Alexa: Sorry I did not understand that.Please try again.")
                    # buttonClick()

#else :
    #time.sleep(2)

entry1.bind('<Return>',get) #to search when we press enter


MyButton1 = ttk.Button(root,text='search',width=10,command=callback)
MyButton1.place(x=260,y=97)

MyButton2 = Button(root,text='Initate',height=0,width=7,font=("Arial bold",9),command=buttonClick)
MyButton2.place(x=268,y=65)

entry1.focus()  #to starting typing with clicking in the bar

root.wm_attributes('-topmost',1)
root.mainloop()
