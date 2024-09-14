import pystray
from PIL import Image, ImageDraw
import speech_recognition as sr
import pyttsx3
import wikipedia
import pywhatkit
import webbrowser
import pyperclip
import pyautogui
import openai
import threading
import time

import os  # For file handling and command generation

openai.api_key = 'sk-hcMtJLZCtMkCU61hiU2ItPAUShGt6xi3E0xIu2IDHlT3BlbkFJqn9GeJIf0w7mTCY22XH6_46ITtToqb_Vb0o7tfuOAA'  # Replace with your valid OpenAI API key

recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Flag to control if the assistant is paused
is_paused = False

# Dictionary of applications to open
applications = {
    "notepad": "C:/Windows/System32/notepad.exe",
    "calculator": "C:/Windows/System32/calc.exe",
    "word": "C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE",
    "excel": "C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE",
    "telegram": "C:/Users/rohit/Downloads/tg/Telegram.exe",
}

# Helper functions
def create_image():
    image = Image.new('RGB', (64, 64), color=(0, 0, 0))
    dc = ImageDraw.Draw(image)
    dc.rectangle((32, 32, 64, 64), fill=(255, 255, 255))
    return image

def speak(text):
    engine.say(text)
    engine.runAndWait()

def take_command():
    global is_paused
    try:
        with sr.Microphone() as source:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)
            query = recognizer.recognize_google(audio)
            print(f"You said: {query}")
            return query.lower()
    except sr.UnknownValueError:
        speak(" ")
        return None
    except sr.RequestError:
        speak("Sorry, there was an error with the request.")
        return None

def open_chatgpt_and_paste():
    webbrowser.open("https://chat.openai.com/")
    time.sleep(5)
    copied_text = pyperclip.paste()
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

def open_gpt():
    webbrowser.open("https://chat.openai.com/")
    time.sleep(5)

def google_search(query):
    webbrowser.open(f"https://www.google.com/search?q={query}")

def process_command(query):
    global is_paused
    if "jarvis stop" in query:
        if not is_paused:
            speak("Voice assistant paused.")
            is_paused = True
        return True  # Don't exit, just pause
    if "i need assistance" in query:
        speak("Don't worry, I will take care of it sirr")
        open_gpt()
    if "resume" in query:
        if is_paused:
            speak("Voice assistant resumed.")
            is_paused = False
        return True  # Resume listening

    if is_paused:
        return True  # If paused, ignore all other commands

    if "exit" in query:
        speak("Goodbye!")
        return False
    elif "hello" in query:
        speak("Hello! How can I help you?")

    elif "play" in query:
        song = query.replace("play", "")
        speak(f"Playing {song}")
        pywhatkit.playonyt(song)
    elif "jarvis help me" in query:
        speak("Opening ChatGPT and submitting your query.")
        open_chatgpt_and_paste()
    elif "search" in query:
        query = query.replace("search", "")
        google_search(query)
    elif any(app in query for app in applications):
        app_name = next(app for app in applications if app in query)
        speak(f"Opening {app_name}.")
        os.startfile(applications[app_name])
    else:

            speak(" ")
    return True

def listen():
    while True:
        query = take_command()
        if query:
            if not process_command(query):
                break

def on_activate(icon, item):
    threading.Thread(target=listen).start()

def on_quit(icon, item):
    global running
    running = False  # Stop the listening loop
    icon.stop()  # Stop the tray icon

# Initialize and run system tray
icon = pystray.Icon("voice_assistant", create_image(), "Voice Assistant")
icon.menu = pystray.Menu(
    pystray.MenuItem('Activate', on_activate),
    pystray.MenuItem('Quit', on_quit)
)

# Start tray
icon.run()
