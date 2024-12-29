# Author: Joseph D. Peralta
# I am going to create a Virtual Assistant for my laptop
# to automate some of the basic redandant task I do everyday
# I installed this <pip install pyttsx3 SpeechRecognition>

# import the libraries we need
import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import subprocess
import pywhatkit
import random
import time
import screen_brightness_control as sbc
import pycaw
import pyjokes
import nltk
from datetime import datetime
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from chatterbot import ChatBot
from chatterbot.trainers import ChatterBotCorpusTrainer
from chatterbot.trainers import ListTrainer
import pyautogui
from time import sleep

# for chat
nltk.download('punkt_tab')

# Initialize the object recognizer and the TTS engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()
print("Initializing chatbot...")
print("Recognizer:", recognizer)
print("Engine:", engine)

# Set the voice to female
voices = engine.getProperty('voices')

# Print available voices and their details
# for voice in voices:
#     print(f"Voice: {voice.name}")
#     print(f" - ID: {voice.id}")
#     print(f" - Gender: {voice.gender}")
#     print(f" - Language: {voice.languages}")

for voice in voices:
    # Set speech rate (default is around 200)
    engine.setProperty('rate', 210)
    # Set volume level (between 0.0 and 1.0)
    engine.setProperty('volume', 1.0)  # Full volume

    if "zira" in voice.name.lower():
        engine.setProperty('voice', voice.id)
        break # select the first available female voice

# Initialize ChatterBot
chatbot = ChatBot('Karen AI')
##trainer = ChatterBotCorpusTrainer(chatbot)
trainer = ListTrainer(chatbot)

## Train the ChatterBot with English corpus data
## trainer.train(
##     'chatterbot.corpus.english.greetings',
##     'chatterbot.corpus.english.conversations'
##     )

# Church QNA
trainer.train([
        "are latter day saints christian",
        "Yes. The Church of Jesus Christ of Latter-day Saints is a Christian church but is neither Catholic nor Protestant. Rather, it is a restoration of the Church of Jesus Christ as originally established by the Savior in the New Testament of the Bible. Latter-day Saints believe God sent His Son, Jesus Christ, to save all mankind from death and their individual sins. Jesus Christ is central to the lives of Church members. They seek to follow His example by being baptized (see Matthew 3:13-17), praying in His holy name (see Matthew 6:9-13), partaking of the sacrament (see Luke 22:19-20), doing good to others (see Acts 10:38) and bearing witness of Him through both word and deed (see James 2:26). The only way to salvation is through faith in Jesus Christ.",
        "what do latter day saints believe about God",
        "God is often referred to in The Church of Jesus Christ of Latter-day Saints as our Heavenly Father because He is the Father of all human spirits and they are created in His image (see Genesis 1:27). It is an appropriate term for God who is kind and just, all wise and all powerful. God the Father, His Son, Jesus Christ, and the Holy Ghost constitute the Godhead or Trinity for Latter-day Saints. Latter-day Saints believe God is embodied, though His body is perfect and glorified.",
        "do latter day saints believe in the trinity",
        "Latter-day Saints most commonly use the term “Godhead” to refer to the Trinity. The first article of faith for the Latter-day Saints reads: “We believe in God, the Eternal Father, and in His Son, Jesus Christ, and in the Holy Ghost.” Latter-day Saints believe God the Father, Jesus Christ and the Holy Ghost are one in will and purpose but are not literally the same being or substance, as conceptions of the Holy Trinity commonly imply.",
        "what is the latter day saint view of the purpose of life",
        "For Latter-day Saints, mortal existence is seen in the context of a great sweep of history, from a pre-earth life where the spirits of all mankind lived with Heavenly Father to a future life in His presence where continued growth, learning and improving will take place. Life on earth is regarded as a temporary state in which men and women are tried and tested — and where they gain experiences obtainable nowhere else. God knew humans would make mistakes, so He provided a Savior, Jesus Christ, who would take upon Himself the sins of the world. To members of the Church, physical death on earth is not an end but the beginning of the next step in God’s plan for His children.",
        "do latter day saints believe in the bible?",
        "Yes. The Church reveres the Bible as the word of God, a sacred volume of scripture. Latter-day Saints cherish its teachings and engage in a lifelong study of its divine wisdom. Moreover, during worship services the Bible is pondered and discussed. Additional books of scripture — including the Book of Mormon— strengthen and reinforce God’s teachings through additional witnesses and provide moving accounts of the personal experiences many individuals had with Jesus Christ. According to Church apostle M. Russell Ballard, “The Book of Mormon does not dilute nor diminish nor deemphasize the Bible. On the contrary, it expands, extends, and exalts it.",
        "what is the book of mormon?",
        "In addition to the Old and New Testaments of the Bible, the Book of Mormon is another testament of Jesus Christ. It contains the writings of ancient prophets, giving an account of God’s dealings with the peoples on the American continent. For Latter-day Saints it stands alongside the Old and New Testaments of the Bible as holy scripture.",
        "what is a latter day saint temple?",
        "Temples existed throughout Biblical times. These buildings were considered the house of the Lord (see 2 Chronicles 2:1-5). Latter-day Saint temples are likewise considered houses of the Lord by Church members. To Latter-day Saints, temples are sacred buildings in which they are taught about the central role of Christ in God’s plan of salvation and their personal relationship with God. In temples, members of the Church make covenants with God to live a virtuous and faithful life. They also offer sacraments on behalf of their deceased ancestors. Latter-day Saint temples are also used to perform marriage ceremonies that promise the faithful eternal life with their families. For members of the Church family is of central importance.",
        "do latter day saints believe in modern day prophets",
        "Yes. The Church is governed today by apostles, reflecting the way Jesus organized His Church in biblical times. Three apostles constitute the First Presidency (consisting of the president or prophet of the Church and his two counselors), and, together with the Quorum of the Twelve Apostles, they have responsibility for leading the Church worldwide and serving as special witnesses of the Lord Jesus Christ. Each is accepted by Church members in a prophetic role corresponding to the apostles in the Bible.",
        "do latter day saints believe that the apostles receive revelations from god",
        "Yes. When Latter-day Saints speak to God, they call it prayer. When God responds through the influence of the Holy Spirit, members refer to this as revelation. Revelation, in its broad meaning, is divine guidance or inspiration; it is the communication of truth and knowledge from God to His children on earth, suited to their language and understanding. It simply means to uncover something not yet known. The Bible illustrates different types of revelation, ranging from dramatic visions to gentle feelings — from the “burning bush” to the “still, small voice.” Latter-day Saints generally believe that divine guidance comes quietly, taking the form of impressions, thoughts and feelings carried by the Spirit of God. Most often, revelation unfolds as an ongoing, prayerful dialogue with God: A problem arises, its dimensions are studied out, a question is asked, and if we have sufficient faith, God leads us to answers, either partial or full. Though ultimately a spiritual experience, revelation also requires careful thought. God does not simply hand down information. He expects us to figure things out through prayerful searching and sound thinking. The First Presidency (consisting of the president or prophet of the Church and his two counselors) and members of the Quorum of the Twelve Apostles receive inspiration to guide the Church as a whole. Individuals are also inspired with revelation regarding how to conduct their lives and help serve others.",
        "do latter day saint women lead in the church",
        "Yes. All women are daughters of a loving Heavenly Father. Women and men are equal in the sight of God. The Bible says, “There is neither Jew nor Greek, there is neither bond nor free, there is neither male nor female: for ye are all one in Christ Jesus” (Galatians 3:28). In the family, a wife and a husband form an equal partnership in leading and raising a family. From the beginning of The Church of Jesus Christ of Latter-day Saints women have played an integral role in the work of the Church. While worthy men hold the priesthood, worthy women serve as leaders, counselors, missionaries, teachers, and in many other responsibilities— they routinely preach from the pulpit and lead congregational prayers in worship services. They serve both in the Church and in their local communities and contribute to the world as leaders in a variety of professions. Their vital and unique contribution to raising children is considered an important responsibility and a special privilege of equal importance to priesthood responsibilities.",
        "do latter day saints believe they can become gods",
        "Latter day Saints believe that is God’s purpose to exalt us to become like Him. But this teaching is often misrepresented by those who caricature the faith. The Latter-day Saint belief is no different than the biblical teaching, which states, “The Spirit itself beareth witness with our spirit, that we are the children of God: and if children, then heirs; heirs of God, and joint-heirs with Christ; if so be that we suffer with him, that we may be also glorified together” (Romans 8:16-17).",
        "do latter day saints believe that they will get their own planet",
        "No. This idea is not taught in Latter-day Saint scripture, nor is it a doctrine of the Church. This misunderstanding stems from speculative comments unreflective of scriptural doctrine. Latter-day Saints believe that we are all sons and daughters of God and that all of us have the potential to grow during and after this life to become like our Heavenly Father (see Romans 8:16-17). The Church does not and has never purported to fully understand the specifics of Christ’s statement that “in my Father’s house are many mansions” (John 14:2).",
        "do some latter day saints wear temple garments",
        "Yes. In our world of diverse religious observance, many people of faith wear special clothing as a reminder of sacred beliefs and commitments. This has been a common practice throughout history. Today, faithful adult members of The Church of Jesus Christ of Latter-day Saints wear temple garments. These garments are simple, white underclothing composed of two pieces: a top piece similar to a T-shirt and a bottom piece similar to shorts. Not unlike the Jewish tallit katan (prayer shawl), these garments are worn underneath regular clothes. Temple garments serve as a personal reminder of covenants made with God to lead good, honorable, Christlike lives. The wearing of temple garments is an outward expression of an inward commitment to follow the Savior. Biblical scripture contains many references to the wearing of special garments. In the Old Testament the Israelites are specifically instructed to turn their garments into personal reminders of their covenants with God (see Numbers 15:37-41). Indeed, for some, religious clothing has always been an important part of integrating worship with daily living. Such practices resonate with Latter-day Saints today. Because of the personal and religious nature of the temple garment, the Church asks all media to report on the subject with respect, treating Latter-day Saint temple garments as they would religious vestments of other faiths. Ridiculing or making light of sacred clothing is highly offensive to Latter-day Saints.",
        "do latter day saints practice polygamy",
        "No. There are over 14 million members of The Church of Jesus Christ of Latter-day Saints and not one of them is a polygamist. The practice of polygamy is strictly prohibited in the Church. The general standard of marriage in the Church has always been monogamy, as indicated in the Book of Mormon (see Jacob 2:27). For periods in the Bible polygamy was practiced by the patriarchs Abraham, Isaac and Jacob, as well as kings David and Solomon. It was again practiced by a minority of Latter-day Saints in the early years of the Church. Polygamy was officially discontinued in eighteen ninety — 122 years ago. Those who practice polygamy today have nothing whatsoever to do with the Church.",
        "what is the position of the Church regarding race relations",
        "The gospel of Jesus Christ is for everyone. The Book of Mormon states, “Black and white, bond and free, male and female; … all are alike unto God” (2 Nephi 26:33). This is the Church’s official teaching. People of all races have always been welcomed and baptized into the Church since its beginning. In fact, by the end of his life in 1844 Joseph Smith, the founding prophet of The Church of Jesus Christ of Latter-day Saints, opposed slavery. During this time some black males were ordained to the priesthood. At some point the Church stopped ordaining male members of African descent, although there were a few exceptions. It is not known precisely why, how or when this restriction began in the Church, but it has ended. Church leaders sought divine guidance regarding the issue and more than three decades ago extended the priesthood to all worthy male members. The Church immediately began ordaining members to priesthood offices wherever they attended throughout the world. The Church unequivocally condemns racism, including any and all past racism by individuals both inside and outside the Church. In 2006, then Church president Gordon B. Hinckley declared that “no man who makes disparaging remarks concerning those of another race can consider himself a true disciple of Christ. Nor can he consider himself to be in harmony with the teachings of the Church. Let us all recognize that each of us is a son or daughter of our Father in Heaven, who loves all of His children.",
        "do latter day saints believe that the garden of eden is in missouri",
        "We do not know exactly where the original site of the Garden of Eden is. While not an important or foundational doctrine, Joseph Smith established a settlement in Daviess County, Missouri, and taught that the Garden of Eden was somewhere in that area. Like knowing the precise number of animals on Noah’s ark, knowing the precise location of the Garden of Eden is far less important to one’s salvation than believing in the Atonement of Jesus Christ.",
        "why do you baptize for the dead",
        "Jesus Christ taught that “except a man be born of water and of the Spirit, he cannot enter into the kingdom of God” (John 3:5). For those who have passed on without the ordinance of baptism, proxy baptism for the deceased is a free will offering. According to Church doctrine, a departed soul in the afterlife is completely free to accept or reject such a baptism — the offering is freely given and must be freely received. The ordinance does not force deceased persons to become members of The Church of Jesus Christ of Latter-day Saints, nor does the Church list deceased persons as members of the Church. In short, there is no change in the religion or heritage of the recipient or of the recipient's descendants — the notion of coerced conversion is utterly contrary to Church doctrine. Of course, proxy baptism for the deceased is nothing new. It was mentioned by Paul in the New Testament (see 1 Corinthians 15:29) and was practiced by groups of early Christians. As part of a restoration of New Testament Christianity, Latter-day Saints continue this practice. All Church members are instructed to submit names for proxy baptism only for their own deceased relatives as an offering of familial love.",
        "why does the church send out missionaries",
        "The missionary effort of The Church of Jesus Christ of Latter-day Saints is based on the New Testament pattern of missionaries serving in pairs, teaching the gospel and baptizing believers in the name of Jesus Christ (see, for example, the work of Peter and John in the book of Acts). More than 52,000 missionaries, most of whom are under the age of 25, are serving missions for the Church at any one time. Missionary work is voluntary, with most missionaries funding their own missions. They receive their assignment from Church headquarters and are sent only to countries where governments allow the Church to operate. In some parts of the world, missionaries are sent only to serve humanitarian or other specialized missions.",
        "why dont latter day saints smoke or drink alcohol",
        "The health code for Latter-day Saints is based on a teaching regarding foods that are healthy and substances that are not good for the human body. Accordingly, alcohol, tobacco, tea, coffee and illegal drugs are forbidden. A 14-year UCLA study, completed in 1997, tracked mortality rates and health practices of 10,000 members of The Church of Jesus Christ of Latter-day Saints in California, indicating that Church members who adhered to the health code had one of the lowest death rates from cancer and cardiovascular disease in the United States. It also found that Church members who followed the code had a life expectancy eight to 11 years longer than the general white population of the United States."
])

def run_assistant():
    speak("Hello! My name is Karen. How can I help you today?")

    # loop to keep it running
    while True:
        command = listen()
        if command:
            process_command(command)

# Function to process our command
def process_command(command):
    
    # Write in notepad
    if 'write something' in command or 'dictation mode' in command:
        if command == 'write something':
            speak("What would you like to write?")
            text_to_write = listen()

            if text_to_write:
                write_to_notepad(text_to_write)
            else: 
                speak("I didn't hear anything to write.")
        else:
            dictation_mode()

    # task responses
    elif 'open youtube' in command:
        speak("Opening Youtube")
        webbrowser.open("https://www.youtube.com")
    elif 'open google' in command:
        speak("Opening Google")
        webbrowser.open("https://www.google.com/")
    elif 'open gmail' in command:
        speak("Opening your email")
        webbrowser.open("https://www.gmail.com")
    elif 'what time is it' in command:
        current_time = datetime.now().strftime('%I:%M')
        speak(f"The time is {current_time}")
    
    # OPEN APPS
    # elif '' in command:
    #     speak("Opening ")
    #     subprocess.Popen(["start", ""], shell=True)
    elif 'open vs code' in command:
        speak("Opening VS Code")
        subprocess.Popen('Code', shell=True)
    elif 'open notepad' in command:
        open_notepad()
    elif 'open word' in command:
        speak("Opening Microsoft Word")
        subprocess.Popen(["start", "winword"], shell=True)
    elif 'open powerpoint' in command:
        speak("Opening Microsoft Powerpoint")
        subprocess.Popen(["start", "powerpnt"], shell=True)
    elif 'open excel' in command:
        speak("Opening Microsoft Excel")
        subprocess.Popen(["start", "excel"], shell=True)
    elif 'open solo leveling' in command:
        speak("Opening Solo Leveling")
        subprocess.Popen(["start", "Netmarble Launcher.exe"], shell=True)
    elif 'open task manager' in command:
        speak("Opening Task Manager")
        subprocess.Popen(["start", "taskmgr.exe"], shell=True)
    elif 'open paint' in command:
        speak("Opening Paint")
        subprocess.Popen(["start", "mspaint"], shell=True)
    elif 'open calculator' in command:
        speak("Opening Calculator")
        subprocess.Popen(["start", "calc"], shell=True)  
    elif 'open snipping tool' in command:
        speak("Opening Snipping Tool")
        subprocess.Popen(["start", "snippingtool"], shell=True)         
    elif 'open chat gpt' in command:
        speak("Opening ChatGPT")
        webbrowser.open("https://chat.openai.com/")
    elif 'open cmd' in command or 'open command prompt' in command:
        speak("Opening Command Prompt")
        subprocess.Popen(["start", "cmd"], shell=True)
    elif 'open media player' in command:
        speak("Opening Media Player")
        subprocess.Popen(["start", "WMPlayer.exe"], shell=True)
    elif 'open file explorer' in command:
        speak("Opening File Explorer")
        subprocess.Popen(["start", "explorer"], shell=True)

    # search or play a song on youtube 
    elif 'search for' in command and 'on youtube' in command:
        # Extract the song name
        song_name = command.replace('search for', '').replace('on youtube', '').strip()
        if song_name:
            speak(f"Searching for {song_name} on YouTube")
            # Format the search query to replace spaces with '+'
            search_query = song_name.replace(' ','+')
            # Generate the YouTube URL
            youtube_url = f"https://www.youtube.com/results?search_query={search_query}"
            # Open youtube search
            webbrowser.open(youtube_url)
        else:
            speak("What song do you want me to search?")
    elif 'play' in command and 'on youtube' in command:
        song_name = command.replace('search for', '').replace('on youtube', '').strip()
        if song_name:
            speak(f"Playing {song_name} on YouTube")
            pywhatkit.playonyt(song_name) # play the first result
        else: 
            speak("Please tell me what song you want to play.")
    elif 'play local music' in command:
        play_local_music()

    # Search something on google
    elif 'search for' in command and 'on google' in command:
        search_query = command.replace('search for', '').replace('on google', '').strip()
        search_on_google(search_query)

    # Task
    elif 'close' in command:
        app_name = command.replace('close', '').strip()
        close_app(app_name)
    elif 'set an alarm for' in command:
        # Extract the time and whether it's AM or PM
        words = command.split()
        try:
            time_part = words[4]  # Example: "7:30"
            am_pm = words[5]  # Example: "AM" or "PM"
            am_pm = am_pm.lower()  # Convert to lowercase for consistency
            
            # Convert time to 24-hour format
            alarm_time = convert_to_24_hour(time_part, am_pm)
            print(f"Alarm time (24-hour format): {alarm_time}")
            set_alarm(alarm_time)
        except IndexError:
            speak("Please specify the time correctly, for example, 'Set an alarm for 7:30 AM'.")
    elif 'set brightness to' in command:
        level = command.replace('set brightness to', '').strip()
        if level == "max":
            level = float(100.0)
        elif level == "5" or level == "five":
            level = float(5.0)
        elif level == "10" or level == "ten":
            level = float(10.0)
        elif level == "20" or level == "twenty":
            level = float(20.0)
        elif level == "50" or level == "fifty":
            level = float(50.0)
        elif level == "75" or level == 'seventy five':
            level = float(75.0)
        else: 
            level = float(1.0)
        set_brightness(level)
    elif 'set volume to' in command:
        vol = command.replace('set volume to', '').strip()
        if vol == "max":
            vol = float(1.0)
        elif vol == "mute":
            vol = float(0.0)
        elif vol == "5" or vol == "five":
            vol = float(0.05)
        elif vol == "10" or vol == "ten":
            vol = float(0.1)
        elif vol == "20" or vol == "twenty":
            vol = float(0.2)
        elif vol == "50" or vol == "fifty":
            vol = float(0.5)
        elif vol == "75" or vol == "seventy five":
            vol = float(0.75)
        else: 
            vol = float(0.2)
        set_volume(vol)
    # Tell a joke
    elif 'tell a joke' in command or 'tell me a joke' in command:
        tell_joke()
    # conversational responses
    elif 'nice' in command:
        speak('Thank you!')
    elif 'karen' in command:
        speak('Yes babe, how can I help you?')   
    elif 'thank you' in command:
        speak("You're welcome!")
    elif 'exit' in command:
        speak("Goodbye!")
        exit()
    elif 'go to sleep' in command:
        speak("I am going to sleep now.")
        exit()
    # else:
    #     speak("Sorry, I cant help you with that.")
    # Chatbot
    else:
        response = get_chatterbot_response(command)
        speak(response)

# Function to make karen(our virtual assistant) speak
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen to our voice
def listen():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

        try:
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
            return command.lower()
        except sr.UnknownValueError:
            speak("Sorry, I did not understand that.")
            return ""
        except sr.RequestError:
            speak("Sorry, my speech service is down.")
            return ""

# Play random music from music folder
def play_local_music():
    music_folder = r"C:/Users/josep/Music/songs"
    songs = os.listdir(music_folder)
    song = random.choice(songs)
    speak(f"Playing {song}")
    subprocess.Popen([os.path.join(music_folder, song)], shell=True)

# Function to set an alarm
def set_alarm(alarm_time):
    speak(f"Setting an alarm for {alarm_time}")
    while True:
        current_time = datetime.now().strftime('%H:%M')
        print(f"Current time: {current_time}")
        if current_time == alarm_time:
            speak("Time to wake up!")
            break
        time.sleep(60)  # Check every minute

# Function to convert 12-hour format to 24-hour format
def convert_to_24_hour(time_str, am_pm):
    hours, minutes = map(int, time_str.split(":"))
    
    if am_pm == "p.m." and hours != 12:
        hours += 12
    elif am_pm == "a.m." and hours == 12:
        hours = 0  # Midnight case
    
    return f"{hours:02}:{minutes:02}"

# Adjust screen brightness
def set_brightness(level):
    sbc.set_brightness(level)
    speak(f"Screen brightness set to {level}%")

# Adjust volume
def set_volume(level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    volume.GetMute()
    volume.GetMasterVolumeLevel()
    volume.GetVolumeRange()
    # Set the master volume level
    volume.SetMasterVolumeLevelScalar(level, None)  # Adjust volume 0.0 - 1.0

# Function to tell random programming jokes
def tell_joke():
    speak(pyjokes.get_joke())

# Search things on goole
def search_on_google(search_query):
    speak(f"Searching {search_query} on google")
    webbrowser.open(f"https://www.google.com/search?q={search_query.replace(' ','+')}")

# Close application
import subprocess

def close_app(app_name):
    try:
        # Use taskkill to close application by executable name
        if app_name == "task manager":
            app_name = "taskmgr"
        elif app_name == "sqlite" or app_name == "sq light":
            app_name = "DB Browser for SQLite"
        elif app_name == "media player":
            app_name = "Microsoft.Media.Player"
        elif app_name == "ms teams" or app_name == "ms themes":
            app_name = "ms-teams"
        subprocess.call(f"taskkill /f /im {app_name}.exe", shell=True)
        speak(f"Closing {app_name}")
    except Exception as e:
        speak(f"Failed to close {app_name}: {str(e)}")

# Function to get a conversational response from ChatterBot
def get_chatterbot_response(user_input):
    response = chatbot.get_response(user_input)
    return str(response)

# Function to type the spoken text into the notepad using pyautogui
def write_to_notepad(text):
    speak(f"Writing '{text}'.")
    
    # split the input so we can handle commands
    words = text.split()

    # check commands
    for word in words:
        if word.lower() == "enter" or word.lower() == "new line":
            pyautogui.press('enter') # simulate pressing enter in keyboard
        else:
            pyautogui.write(word + " ", interval=0.3) # typing speed

# Function to open notepad
def open_notepad():
    speak("Opening notepad")
    # use .Popen instead of .run(this wait for the process to close)
    # detach the process on windows using creationFlags
    subprocess.Popen("notepad.exe", creationflags=subprocess.CREATE_NEW_CONSOLE)
    time.sleep(2) # wait for the notepad to open

# Function that handle dictation mode
def dictation_mode():
    speak("Dictation mode started. Say 'stop dictation' to stop.")
    while True:
        text_to_write = listen()

        # exit dictation mode when the user says "stop dictation"
        if 'stop dictation' in text_to_write:
            speak("Exiting dictation mode.")
            break

        # write the spoken words
        write_to_notepad(text_to_write + " ") # add space after every input

run_assistant()