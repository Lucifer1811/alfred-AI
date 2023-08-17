import win32com.client
import speech_recognition as sr
import os
import webbrowser
import datetime
import openai
from key_oAI import api_key


speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(query):
    speaker.Speak(query)

chatStr = ""

def chat(query):
    global chatStr
    openai.api_key = api_key
    chatStr += f"Wayne: {query}\n Alfred:"
    print(chatStr)
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    say(response.choices[0].text)
    chatStr += f"{response.choices[0].text}\n"
    

def ai(prompt):
    openai.api_key = api_key
    text = f"OpenAI response for: {prompt}\n"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        text += response.choices[0].text
        if not os.path.exists("ai_responses"):
            os.mkdir("ai_responses")
        with open(f"ai_responses/{''.join(prompt.split('AI')[1:]).strip()}.txt", "w") as f:
            f.write(text)
    except Exception as exeption: 
        return "Sorry, there is some error on my end."
    
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occured. Sorry"
        
if __name__ == '__main__':
    print('Hello I am Alfred A.I')
    say("Hello I am Alfred in the form of AI")
    while True:
        print("Listening...")
        query = takeCommand()
        ends = ["bye", "exit", "quit", "stop", "fuck off", "shut up", "shut down", "kill yourself"]
        for end in ends:
            if end.lower() in query.lower():
                exit()    
            
        sites = [["youtube", "https://youtube.com"],["wikipedia", "https://wikipedia.com"],["netflix", "https://netflix.com"],["prime video", "https://primevideo.com"],
                 ["hotstar", "https://hotstar.com"],["music YT", "https://music.youtube.com"],["spotify", "https://spotify.com"],["udemy", "https://udemy.com"],]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} Sir...")
                webbrowser.open(site[1])
        apps = [["discord","C:/Users/lucif/AppData/Local/Discord/app-1.0.9016/Discord.exe"],
                ["WhatsApp","C:/Users/lucif/OneDrive/Desktop/WhatsApp.lnk"]]
        for app in apps:
            if f"open {app[0]}".lower() in query.lower():
                say(f"Opening {app[0]} Sir...")
                os.startfile(app[1])
                
        if "time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir the time is {strfTime}")
            
        if "Sir Alfred AI".lower() in query.lower():
            ai(prompt = query)
            say("Yes Sir")
            
        if "reset chat" in query.lower():
            chatStr = ""
        
        else:
            print("Chatting...")
            chat(query)
        
        
        # say(query)