from email.mime.multipart import MIMEMultipart

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
import subprocess
from email.mime.base import MIMEBase
from email import encoders
import re
import os
import argparse
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import json

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
import base64
import os

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Path to the credentials file
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'

import sys

def get_user_input(prompt):
    sys.stdout.write(prompt)
    sys.stdout.flush()
    return sys.stdin.readline().strip()

def authenticate_gmail_api():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds


def create_message(sender, to, subject, message_text, file_path=None):
    """Create a MIMEText email message with optional file attachment."""
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg_body = MIMEText(message_text, 'plain')
    message.attach(msg_body)

    if file_path:
        if os.path.isfile(file_path):
            try:
                part = MIMEBase('application', 'octet-stream')
                with open(file_path, 'rb') as f:
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
                message.attach(part)
            except Exception as e:
                print(f"Failed to attach file: {e}")
        else:
            print(f"File does not exist: {file_path}")

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}


import tkinter as tk
from tkinter import simpledialog, messagebox

def get_email_details():
    root = tk.Tk()
    root.title("Email Input Form")  # Title of the window
    root.geometry("400x300")  # Width x Height

    # Create a frame to organize the widgets
    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

    # Create and place labels and entry fields
    tk.Label(frame, text="Recipient Email:", font=('Arial', 12)).grid(row=0, column=0, sticky=tk.W, pady=5)
    recipient_var = tk.StringVar()
    tk.Entry(frame, textvariable=recipient_var, width=50, font=('Arial', 12)).grid(row=0, column=1, pady=5)

    tk.Label(frame, text="Subject:", font=('Arial', 12)).grid(row=1, column=0, sticky=tk.W, pady=5)
    subject_var = tk.StringVar()
    tk.Entry(frame, textvariable=subject_var, width=50, font=('Arial', 12)).grid(row=1, column=1, pady=5)

    tk.Label(frame, text="Body:", font=('Arial', 12)).grid(row=2, column=0, sticky=tk.W, pady=5)
    body_var = tk.StringVar()
    tk.Entry(frame, textvariable=body_var, width=50, font=('Arial', 12)).grid(row=2, column=1, pady=5)

    tk.Label(frame, text="File Path (optional):", font=('Arial', 12)).grid(row=3, column=0, sticky=tk.W, pady=5)
    file_path_var = tk.StringVar()
    tk.Entry(frame, textvariable=file_path_var, width=50, font=('Arial', 12)).grid(row=3, column=1, pady=5)

    def on_submit():
        recipient = recipient_var.get()
        subject = subject_var.get()
        body = body_var.get()
        file_path = file_path_var.get()

        if not recipient or not subject or not body:
            messagebox.showerror("Input Error", "Recipient, Subject, and Body are required fields.")
            return

        # Call the function to send email
        send_email_with_gmail(recipient, subject, body, file_path)
        messagebox.showinfo("Success", "Email sent successfully!")
        root.destroy()  # Close the window

    submit_button = tk.Button(frame, text="Send Email", command=on_submit, font=('Arial', 12))
    submit_button.grid(row=4, column=1, pady=20)

    root.mainloop()

def send_email_with_gmail(recipient, subject, body, file_path=None):
    try:
        # Authenticate Gmail API
        creds = authenticate_gmail_api()
        service = build('gmail', 'v1', credentials=creds)

        # Validate email
        if not recipient or not re.match(r"[^@]+@[^@]+\.[^@]+", recipient):
            print("That doesn't seem like a valid email address. Please try again.")
            return

        # Create email
        sender = 'rohithpatil006@gmail.com'  # Replace with your Gmail
        message = create_message(sender, recipient, subject, body, file_path)

        # Send email
        message = (service.users().messages().send(userId="me", body=message).execute())
        print(f"Email sent to {recipient}.")

    except HttpError as error:
        print(f'An error occurred: {error}')
        print("Failed to send the email.")
    except Exception as e:
        print(f"Error: {e}")
        print("An error occurred while sending the email.")






def enable_bluetooth():
    try:
        # Check if Bluetooth service is already running
        check_service = subprocess.run(
            "powershell.exe Get-Service -Name bthserv",
            capture_output=True, text=True
        )
        service_status = check_service.stdout

        if "Running" not in service_status:
            # Start Bluetooth service if it's not running
            start_service = subprocess.run(
                "powershell.exe Start-Service -Name bthserv",
                capture_output=True, text=True
            )
            if start_service.returncode == 0:
                print("Bluetooth service started.")
            else:
                print("Failed to start Bluetooth service.")
                return False
        else:
            print("Bluetooth service is already running.")

        # Check if Bluetooth is enabled (alternative method)
        # This assumes that the Bluetooth adapter is present and managed by the OS
        bt_check = subprocess.run(
            "powershell.exe Get-PnpDevice -Class Bluetooth",
            capture_output=True, text=True
        )
        if "Bluetooth" in bt_check.stdout:
            print("Bluetooth adapter found.")
            return True
        else:
            print("Bluetooth adapter not found.")
            return False

    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        return False

def add_bluetooth_device():
    try:
        # Ensure Bluetooth is enabled
        if not enable_bluetooth():
            speak("Failed to enable Bluetooth.")
            return
        # Open Bluetooth settings to add a new device
        subprocess.run("start ms-settings:bluetooth", shell=True, check=True)
        speak("Opening Bluetooth settings to add a new device.")
    except subprocess.CalledProcessError as e:
        speak("Failed to open Bluetooth settings.")
        print(e)


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
    "file explorer":"C:/Users/rohit/OneDrive/Pictures/Documents/Desktop/File Explorer - Shortcut.lnk",
    "downloads":"C:/Users/rohit/Downloads",
    "brave":"C:Program Files/BraveSoftware/Brave-Browser/Application/brave.exe",
}

import requests

def news_headlines():
    url = (
            "https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=1f2d25287eee43f5864f3bad35726e80")

    response = requests.get(url)

    if response.status_code == 200:  # Check if the request was successful
        data = response.json()  # Get the response as JSON
        articles = data.get('articles', [])
        if articles:
            for article in articles:
                title = article.get('title')
                if title:
                    print(title)
                    speak(title)
        else:
            print("No articles found.")
    else:
        print(f"Error: {response.status_code}")

def business_headlines():
    url = (
            "https://newsapi.org/v2/top-headlines/sources?category=businessapiKey=API_KEY")

    response = requests.get(url)

    if response.status_code == 200:  # Check if the request was successful
        data = response.json()  # Get the response as JSON
        articles = data.get('articles', [])
        if articles:
            for article in articles:
                title = article.get('title')
                country=article.get('country')
                if title:
                    print(title,country)

                    speak(title)
        else:
            print("No articles found.")
    else:
        print(f"Error: {response.status_code}")

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
def get_wikipedia_summary(query):
    topic = query.replace("i need info about", "").strip()
    try:
        summary = wikipedia.summary(topic, sentences=5)
        speak(f"According to Wikipedia, {summary}")
    except wikipedia.DisambiguationError as e:
        speak("The topic is ambiguous. Please be more specific.")
    except wikipedia.PageError:
        speak("Sorry, I couldn't find any information on that topic.")
    except Exception as e:
        speak("An error occurred while fetching the information.")

def open_chatgpt_and_paste():
    webbrowser.open("https://chat.openai.com/")
    time.sleep(5)
    copied_text = pyperclip.paste()
    pyautogui.hotkey("ctrl","V")
    pyautogui.press("enter")

def auto_restart():
    speak("confirm it")
    res=take_command()
    if res=="do it":
        pyautogui.hotkey("win", "x")
        pyautogui.press("u")
        pyautogui.press("r")
    else:
        return None

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
    if "turn on bluetooth" in query:
        speak("opening bluetooth")
        add_bluetooth_device()
    if "open chat" in query:
        speak("Okay,sirr")
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

    if "send an email" in query:
        get_email_details()
    elif "connect to bluetooth" in query:
        add_bluetooth_device()
    elif "restart" in query:
        auto_restart()
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
    elif "headlines" in query:
        speak("headlines as of today, are")
        news_headlines()
    elif "i need info about" in query:
        speak("extracting info")
        get_wikipedia_summary(query)

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
