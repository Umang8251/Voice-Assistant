#Python Project Final
import pyttsx3
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import datetime
import subprocess
import keyboard
import pyjokes
#from AppOpener import open
# init pyttsx
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")

engine.setProperty('voice', voices[1].id)  # 1 for female and 0 for male voice


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
  hour=datetime.datetime.now().hour
  if hour >= 0 and hour < 12:
      speak("Good Morning")
  elif hour >= 12 and hour < 18:
      speak("Good Afternoon")
  else:
      speak("Good Evening")


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print("User said:" + query + "\n")
    except Exception as e:
        print(e)
        speak("I didnt understand")
        return "None"
    return query

#def note(text):
   # date = datetime.datetime.now()
    #file_name = str(date).replace(":", "-") + "-note.txt"

    #with open(file_name, "w") as f:
       # f.write(text)

    #subprocess.open(["notepad.exe", file_name])


if __name__ == '__main__':

    wishMe()
    speak("Hello user ")
    speak("How can i help you")

    while True:
        query = take_command().lower()
        if 'wikipedia' in query:
            speak("Searching Wikipedia ...")
            query = query.replace("wikipedia", '')
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            speak(results)

        elif 'are you' in query:
            speak("I am amigo developed by Saanvi,Namita,Sakshi and Umang")

       # elif 'take notes' in query:
           # query = query.replace("make a note", "")
            #
            # note(query)


        elif 'date' in query:
            t=datetime.date.today()
            speak(t)

        elif 'time' in query:
            hour=datetime.datetime.now().time
            speak(hour)

        elif 'open microsoft edge' in query:
            speak("opening microsoft edge")
            webbrowser.open(r"C:\Users\Saanvi\Desktop\Microsoft Edge.lnk") 

        elif 'open calculator' in query:
            speak("opening calculator")
            subprocess.Popen('C:\\Windows\\System32\\calc.exe')

        #elif 'open paint' in query:
            speak("opening paint")
            open("paint")
        
        #elif 'open snipping tool' in query:
            speak("opening snipping tool")
            open("snipping tool")

        #elif 'open settings'  in query:
            speak("opening settings")
            open("settings")

        #elif 'open photos' in query:
            speak("opening photos")
            open("photos")    

        elif 'open youtube' in query:
            speak("opening youtube")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("opening gmail")
            webbrowser.open("google.com")

        elif 'open gmail' in query:
            speak("opening google")
            webbrowser.open("gmail.com")

        elif 'open github' in query:
            speak("opening github")
            webbrowser.open("github.com")

        elif 'open stack overflow' in query:
            speak("opening stackoverflow")
            webbrowser.open("stackoverflow.com")

        elif 'open spotify' in query:
            speak("opening spotify")
            webbrowser.open("spotify.com")

        elif 'open whatsapp' in query:
            speak("opening whatsapp")
            webbrowser.open("whatsapp.com")

        elif 'news' in query:
            news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/city/bangalore")
            speak('Here are some headlines from the Times of India, Happy reading')

        #elif  'corona' in query:
            news = webbrowser.open_new_tab("https://www.worldometers.info/coronavirus/")
            speak('Here are the latest covid-19 numbers')

        elif 'play music' in query:
            speak("opening music")
            webbrowser.open("spotify.com")

       # elif 'weather' in query:
            speak('Here is the weather in bangalore')
            open("weather")

        #elif 'camera' in query:
            speak("opening camera")
            open("camera")
        
        elif 'local disk d' in query:
            speak("opening local disk D")
            webbrowser.open("D://")
        
        elif 'local disk c' in query:
            speak("opening local disk C")
            webbrowser.open("C://")
        
        elif 'local disk e' in query:
            speak("opening local disk E")
            webbrowser.open("E://")
        
        elif 'append file' in query:
            speak("what needs to be appended")
            import aspose.words as aw
            doc = aw.Document("input.doc")
            builder = aw.DocumentBuilder(doc)
            builder.move_to_document_start()
            builder.writeln(input("what is the modification:"))
            doc.save("Input3.doc")
        
        elif "replace text" in query:
            speak("what needs to be replaced")
            s=input("enter text to replace:")
            f=open("umang.txt","r+")
            f.truncate(0)
            f.write(s)
            f.close()
            speak("text successfully replaced")
        
        elif 'delete line' in query:
            speak("which line needs to be deleted")
            def delete_file(filename,line_no):
                with open(filename) as file:
                    lines=file.readlines()
                if (line_no<=len(lines)):
                    del lines[line_no - 1]
                    with open(filename,"w") as file:
                        for line in lines:
                            file.write(line)
                else:
                    print("lines",line_no,"not in file")
                print(lines)
            delete_filename=input('file:')
            delete_lineno=int(input("line:"))
            delete_file(delete_filename,delete_lineno)
            speak("line deleted")
        elif 'addition' in query:
            filename=input("file:")
            h= open(filename, 'r')
            # Reading from the file
            content = h.readlines()
            # Variable for storing the sum
            a = 0
            # Iterating through the content
            # Of the file
            for line in content:
                for i in line:
                    # Checking for the digit in
                    # the string
                    if i.isdigit() == True:
                        a += int(i)
            print("The sum is:", a)
            speak("the sum is printed")
        elif 'replace capital'in query:
            def capital_replace():
                f=open(filename,"r+")
                data=f.read()
                f.seek(0)
                for i in data:
                    if i.isupper():
                        f.write(i.lower())
                    else:
                        f.write(i)
                f.close()
            filename=input("file:")
            capital_replace()
            speak('capital letters successfully replaced')
        
        app_string = ["open word", "open powerpoint", "open excel"]
        app_link = [r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk",r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk", r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" ]
        if app_string[0] in query:
                os.startfile(app_link[0])
                speak("Microsoft office word is opening now")
        elif app_string[1] in query:
                os.startfile( app_link[1])
                speak("Microsoft office powerpoint is opening now")
        elif app_string[2] in query:
                os.startfile(app_link[2])
                speak("Microsoft office excel is opening now")
        
        elif 'joke' in query:
            speak(pyjokes.get_joke())
        
        elif 'sleep' in query:
            speak('Your personal assistant is shutting down, Good bye')
            exit(0)

