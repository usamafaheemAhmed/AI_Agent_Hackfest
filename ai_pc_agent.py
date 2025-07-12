import speech_recognition as sr
import pyttsx3
import pyautogui
import os
import webbrowser

# Initialize recognizer and text-to-speech engine
r = sr.Recognizer()
engine = pyttsx3.init()

def speak(text):
    print("Assistant:", text)
    engine.say(text)
    engine.runAndWait()

def listen():
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
    try:
        command = r.recognize_google(audio)
        print("You said:", command)
        return command.lower()
    except sr.UnknownValueError:
        speak("Sorry, I didn't understand that.")
        return ""
    except sr.RequestError:
        speak("Speech service error.")
        return ""

def execute_command(cmd):
    if cmd.startswith("open "):
        app = cmd.split("open", 1)[1].strip()
        speak(f"Opening {app}")
        os.system(f"start {app}")
        
    elif cmd.startswith("close "):
        process = cmd.split("close", 1)[1].strip()
        speak(f"Closing {process}")
        os.system(f"taskkill /im {process}.exe /f")
        
    elif cmd.startswith("search"):
        query = cmd.split("search", 1)[1].strip()
        speak(f"Searching for {query}")
        webbrowser.open(f"https://www.google.com/search?q={query}")
        
    elif cmd.startswith("type "):
        text = cmd.split("type", 1)[1].strip()
        speak(f"Typing: {text}")
        pyautogui.write(text)
        
    elif cmd in ["exit", "quit", "goodbye"]:
        speak("Goodbye!")
        exit()

    else:
        speak("Trying to execute your command...")
        try:
            os.system(cmd)
        except Exception as e:
            speak(f"Sorry, I couldn't run that. Error: {str(e)}")


# Main loop
speak("Voice control activated. Say a command.")
while True:
    command = listen()
    if command:
        execute_command(command)
