# JARVIS-Empowering-Access-and-Enhancing-Usability-Through-Smart-Voice-Assistance
JARVIS is an AI-powered voice assistant built to enhance digital accessibility and streamline user interaction with the web through intuitive voice commands. Designed with a social impact mindset, JARVIS helps users—especially those from underserved communities—easily discover local job opportunities, access free education, make template resume to.
code for backend:
import speech_recognition as sr
import pyttsx3
import os
import datetime
import subprocess
import sys
import pywhatkit
import wikipedia
import webbrowser
from dotenv import load_dotenv
import openai
import glob
import pathlib
import win32com.client
from fuzzywuzzy import fuzz
import noisereduce as nr
import numpy as np
import urllib.parse
import pyautogui
import time

# Load environment variables from .env file
load_dotenv()

# Set OpenAI API key
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    print("Warning: OPENAI_API_KEY not found in environment variables. Some features may not work.")
else:
    openai.api_key = openai_api_key

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
recognizer = sr.Recognizer()

# Store conversation history for context-aware responses
conversation_history = []

def speak(text):
    print("Jarvis:", text)
    engine.say(text)
    engine.runAndWait()

def get_all_applications():
    """Retrieve a list of all installed applications using Windows Shell."""
    apps = []
    shell = win32com.client.Dispatch("Shell.Application")
    
    folders = [
        shell.NameSpace(0x0),  # Desktop
        shell.NameSpace(0x7),  # Programs
        shell.NameSpace(0x1a),  # User Programs
    ]
    
    program_files = [
        r"C:\Program Files",
        r"C:\Program Files (x86)"
    ]
    
    for folder in folders:
        if folder:
            items = folder.Items()
            for item in items:
                if item.Path.endswith('.lnk') or item.Path.endswith('.exe'):
                    apps.append(item.Path)
    
    for path in program_files:
        for app in glob.glob(os.path.join(path, "**", "*.exe"), recursive=True):
            apps.append(app)
    
    uwp_apps = [
        'xbox://',
        'ms-windows-store://',
        'ms-calculator://',
        'ms-photos://',
        'ms-settings://',
        'ms-edge://',
        'ms-paint://',
        'linkedin://',
        'powerpoint://',
    ]
    apps.extend(uwp_apps)
    
    return apps

def fuzzy_match_command(command, known_commands):
    """Use fuzzy matching to correct misrecognized voice commands."""
    best_match = None
    highest_score = 0
    for known in known_commands:
        score = fuzz.ratio(command.lower(), known.lower())
        if score > highest_score and score > 70:
            highest_score = score
            best_match = known
    return best_match

def fuzzy_match_website(website_name):
    """Use fuzzy matching to correct misrecognized website names."""
    known_websites = [
        'wikipedia', 'gmail', 'youtube', 'srm websites', 'linkedin', 'google maps', 'deepseek',
        'chatgpt', 'openai', 'grok', 'gemini', 'microsoft copilot', 'copilot'
    ]
    matched_website = fuzzy_match_command(website_name, known_websites)
    if matched_website:
        print(f"Fuzzy matched '{website_name}' to '{matched_website}'")
        return matched_website
    return website_name

def reduce_noise(audio_data, sample_rate):
    """Reduce background noise from audio data."""
    return nr.reduce_noise(y=audio_data, sr=sample_rate)

def open_software(software_name):
    try:
        known_commands = [
            'chrome', 'edge', 'notepad', 'calculator', 'paint', 'task manager',
            'command prompt', 'cmd', 'play', 'weather', 'openai', 'chatgpt',
            'grok', 'gemini', 'deepseek', 'google map', 'maps', 'my files',
            'folders', 'xbox', 'photos', 'opera', 'microsoft copilot', 'copilot',
            'linkedin', 'microsoft store', 'powerpoint', 'all apps'
        ]
        
        matched_command = fuzzy_match_command(software_name, known_commands)
        if matched_command:
            speak(f"Hmm, did you mean {matched_command}? Let me open that for you!")
            software_name = matched_command
        
        if 'chrome' in software_name:
            speak('Sure, let’s get Chrome up and running for you!')
            subprocess.Popen(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        elif 'edge' in software_name:
            speak('Alright, opening Microsoft Edge now!')
            subprocess.Popen(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        elif 'notepad' in software_name:
            speak('Opening Notepad—great for jotting down some quick notes!')
            subprocess.Popen(['notepad.exe'])
        elif 'calculator' in software_name:
            speak('Here’s the Calculator—let’s crunch some numbers!')
            subprocess.Popen(['calc.exe'])
        elif 'paint' in software_name:
            speak('Opening Paint—feeling artistic today, huh?')
            subprocess.Popen(['mspaint.exe'])
        elif 'task manager' in software_name:
            speak('Let’s take a look at Task Manager!')
            subprocess.Popen(['Taskmgr.exe'])
        elif 'command prompt' in software_name or 'cmd' in software_name:
            speak('Opening Command Prompt—ready to run some commands!')
            subprocess.Popen(['cmd.exe'])
        elif 'play' in software_name:
            speak('Let’s find something fun to play on YouTube!')
            pywhatkit.playonyt(software_name.replace('play', '').strip())
        elif 'weather' in software_name:
            speak("Let’s check the weather—hope it’s a sunny day!")
            webbrowser.open("https://www.google.com/search?q=current+weather")
        elif 'openai' in software_name or 'chatgpt' in software_name:
            speak('Let’s chat with ChatGPT—should be fun!')
            all_apps = get_all_applications()
            chatgpt_found = False
            for app in all_apps:
                if 'chatgpt' in app.lower() or 'openai' in app.lower():
                    if app.endswith('.lnk') or app.endswith('.exe'):
                        subprocess.Popen(['start', '', app], shell=True)
                        chatgpt_found = True
                        break
                    elif '://' in app:
                        subprocess.Popen(['start', app], shell=True)
                        chatgpt_found = True
                        break
            if not chatgpt_found:
                webbrowser.open('https://chat.openai.com')
        elif 'grok' in software_name:
            speak('Opening Grok by xAI—let’s see what it’s all about!')
            webbrowser.open('https://grok.x.ai')
        elif 'gemini' in software_name:
            speak('Opening Google Gemini—here we go!')
            webbrowser.open('https://gemini.google.com')
        elif 'deepseek' in software_name:
            speak('Opening DeepSeek—let’s dive into some AI fun!')
            webbrowser.open('https://chat.deepseek.com')
        elif 'google map' in software_name or 'maps' in software_name:
            speak('Opening Google Maps—where are we heading today?')
            webbrowser.open('https://www.google.com/maps')
        elif 'my files' in software_name or 'folders' in software_name:
            speak('Let’s open your Documents folder—here you go!')
            documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
            subprocess.Popen(['explorer.exe', documents_path])
        elif 'xbox' in software_name:
            speak('Opening the Xbox app—time for some gaming!')
            subprocess.Popen(['start', 'xbox://'], shell=True)
        elif 'photos' in software_name:
            speak('Opening Photos app—let’s check out your memories!')
            subprocess.Popen(['start', 'ms-photos://'], shell=True)
        elif 'opera' in software_name:
            speak('Opening Opera browser—nice choice!')
            try:
                subprocess.Popen(r"C:\Program Files\Opera\opera.exe")
            except FileNotFoundError:
                speak("Hmm, Opera isn’t in the usual spot—let me look for it!")
                all_apps = get_all_applications()
                for app in all_apps:
                    if 'opera' in app.lower() and app.endswith('.exe'):
                        subprocess.Popen([app])
                        break
                else:
                    speak("Looks like Opera might not be installed—want to try another browser?")
        elif 'microsoft copilot' in software_name or 'copilot' in software_name:
            speak('Opening Microsoft Copilot—let’s get some AI help!')
            all_apps = get_all_applications()
            copilot_found = False
            for app in all_apps:
                if 'copilot' in app.lower():
                    if app.endswith('.lnk') or app.endswith('.exe'):
                        subprocess.Popen(['start', '', app], shell=True)
                        copilot_found = True
                        break
                    elif '://' in app:
                        subprocess.Popen(['start', app], shell=True)
                        copilot_found = True
                        break
            if not copilot_found:
                webbrowser.open('https://copilot.microsoft.com')
        elif 'linkedin' in software_name:
            speak('Opening LinkedIn—time to connect with some folks!')
            all_apps = get_all_applications()
            linkedin_found = False
            for app in all_apps:
                if 'linkedin' in app.lower():
                    if app.endswith('.lnk') or app.endswith('.exe'):
                        subprocess.Popen(['start', '', app], shell=True)
                        linkedin_found = True
                        break
                    elif 'linkedin://' in app:
                        subprocess.Popen(['start', app], shell=True)
                        linkedin_found = True
                        break
            if not linkedin_found:
                webbrowser.open('https://www.linkedin.com')
        elif 'microsoft store' in software_name:
            speak('Opening Microsoft Store—let’s browse some apps!')
            try:
                result = subprocess.run(['powershell', '-Command', 'Get-AppxPackage -Name Microsoft.WindowsStore'], capture_output=True, text=True)
                if 'Microsoft.WindowsStore' in result.stdout:
                    subprocess.run(['powershell', '-Command', 'Start-Process ms-windows-store:'], shell=True, check=True)
                else:
                    speak("Looks like Microsoft Store isn’t installed—maybe check your Windows settings?")
            except Exception as e:
                speak("I couldn’t open the Microsoft Store—it might be disabled or something. What else can I help with?")
                print(f"Error opening Microsoft Store: {e}")
        elif 'powerpoint' in software_name:
            speak('Opening PowerPoint—let’s make some awesome slides!')
            try:
                office_paths = [
                    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
                    r"C:\Program Files\Microsoft Office\root\Office15\POWERPNT.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office15\POWERPNT.EXE"
                ]
                powerpoint_found = False
                for path in office_paths:
                    if os.path.exists(path):
                        subprocess.Popen([path], shell=True)
                        powerpoint_found = True
                        break
                if not powerpoint_found:
                    all_apps = get_all_applications()
                    for app in all_apps:
                        if 'powerpoint' in app.lower():
                            if app.endswith('.lnk') or app.endswith('.exe'):
                                subprocess.Popen(['start', '', app], shell=True)
                                powerpoint_found = True
                                break
                            elif 'powerpoint://' in app:
                                subprocess.Popen(['start', app], shell=True)
                                powerpoint_found = True
                                break
                if not powerpoint_found:
                    speak("I couldn’t find PowerPoint—maybe Microsoft Office isn’t installed?")
            except Exception as e:
                speak("Oops, I couldn’t open PowerPoint—something went wrong!")
                print(f"Error opening PowerPoint: {e}")
        elif 'all apps' in software_name:
            speak('Alright, let’s open all available apps—this might take a sec!')
            all_apps = get_all_applications()
            opened_apps = 0
            for app in all_apps:
                try:
                    if app.endswith('.lnk') or app.endswith('.exe'):
                        subprocess.Popen(['start', '', app], shell=True)
                    elif '://' in app:
                        subprocess.Popen(['start', app], shell=True)
                    opened_apps += 1
                    subprocess.run(['timeout', '1'], shell=True, capture_output=True)
                except Exception as e:
                    print(f"Failed to open {app}: {e}")
            speak(f"I opened {opened_apps} apps for you! Some might not launch due to system restrictions.")
        else:
            speak(f"Hmm, let me see if I can find {software_name} for you!")
            all_apps = get_all_applications()
            found = False
            for app in all_apps:
                if software_name in app.lower() and (app.endswith('.exe') or app.endswith('.lnk') or '://' in app):
                    try:
                        if app.endswith('.lnk') or app.endswith('.exe'):
                            subprocess.Popen(['start', '', app], shell=True)
                        else:
                            subprocess.Popen(['start', app], shell=True)
                        speak(f"Found and opened {software_name}—you’re all set!")
                        found = True
                        break
                    except Exception as e:
                        print(f"Failed to open {app}: {e}")
            if not found:
                speak(f"Hmm, I couldn’t find {software_name} on your system. Maybe try something like ‘open Notepad’ or ‘open Chrome’ instead? What else can I do for you?")
    except Exception as e:
        speak("Oops, something went wrong while trying to open that—let’s try something else! What’s on your mind?")
        print(e)

def open_website_from_google(website_name):
    """Open a website by searching for it on Google."""
    try:
        if not website_name:
            speak("I’d love to help! Which website are we looking for? Maybe something like ‘Wikipedia’ or ‘YouTube’?")
            return
        corrected_website = fuzzy_match_website(website_name)
        speak(f"Let’s search for {corrected_website} on Google—I bet we’ll find it!")
        encoded_website = urllib.parse.quote(corrected_website)
        url = f"https://www.google.com/search?q={encoded_website}"
        print(f"Attempting to open URL: {url}")
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if os.path.exists(chrome_path):
            subprocess.run([chrome_path, url], check=True)
        else:
            edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            if os.path.exists(edge_path):
                subprocess.run([edge_path, url], check=True)
            else:
                speak("I couldn’t find Chrome or Edge on your system. Want to set a default browser and try again?")
                print("Neither Chrome nor Edge found at default paths.")
    except subprocess.CalledProcessError as e:
        speak("I couldn’t open that website—maybe there’s an issue with your browser or internet. Want to try something else, like opening an app?")
        print(f"Subprocess error: {e}")
    except Exception as e:
        speak("Oops, I couldn’t search for that website on Google—let’s try something different! What else can I help you with?")
        print(f"Error searching website on Google: {e}")

def close_software(software_name):
    try:
        tasks = {
            'chrome': 'chrome.exe',
            'edge': 'msedge.exe',
            'notepad': 'notepad.exe',
            'calculator': 'calculator.exe',
            'paint': 'mspaint.exe',
            'task manager': 'taskmgr.exe',
            'command prompt': 'cmd.exe',
            'google': 'chrome.exe',
            'openai': 'chrome.exe',
            'chatgpt': 'chrome.exe',
            'grok': 'chrome.exe',
            'gemini': 'chrome.exe',
            'deepseek': 'chrome.exe',
            'google map': 'chrome.exe',
            'maps': 'chrome.exe',
            'xbox': 'XboxApp.exe',
            'photos': 'Photos.exe',
            'opera': 'opera.exe',
            'microsoft copilot': 'msedge.exe',
            'copilot': 'msedge.exe',
            'linkedin': 'msedge.exe',
            'microsoft store': 'WinStore.App.exe',
            'powerpoint': 'POWERPNT.EXE'
        }

        for key, val in tasks.items():
            if key in software_name:
                speak(f"Closing {key} for you—done! Anything else you’d like to do?")
                os.system(f"taskkill /f /im {val}")
                return
        speak(f"I couldn’t find {software_name} running. Maybe it’s already closed? What else can I help with?")
    except Exception as e:
        speak("Something went wrong while closing that—no worries, let’s try something else! What’s next?")
        print(e)

def close_specific_tab(tab_name):
    """Close a specific tab in the active browser."""
    try:
        if 'chrome' in tab_name or 'google' in tab_name:
            speak(f"Closing the {tab_name} tab for you!")
            subprocess.run(['taskkill', '/IM', 'chrome.exe', '/FI', 'WINDOWTITLE eq ' + tab_name], shell=True)
        elif 'edge' in tab_name:
            speak(f"Closing the {tab_name} tab for you!")
            subprocess.run(['taskkill', '/IM', 'msedge.exe', '/FI', 'WINDOWTITLE eq ' + tab_name], shell=True)
        else:
            speak(f"Hmm, I couldn’t find a {tab_name} tab to close. Make sure the tab is open and try again!")
    except Exception as e:
        speak(f"Oops, I couldn’t close the {tab_name} tab—let’s try something else!")
        print(f"Error closing tab: {e}")

def control_application(command, active_app):
    """Control an active application via voice command."""
    try:
        # YouTube: Play a specific video
        if 'youtube' in active_app and 'play this video' in command:
            speak("Sure, let’s find and play a video for you! What video would you like to watch?")
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=1)
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
                video_query = recognizer.recognize_google(audio, language='en-US').lower()
                speak(f"Playing {video_query} on YouTube now!")
                pywhatkit.playonyt(video_query)

        # Notepad: Type text
        elif 'notepad' in active_app and 'type' in command:
            text_to_type = command.replace('type', '').strip()
            speak(f"Typing '{text_to_type}' in Notepad for you!")
            pyautogui.write(text_to_type)
            pyautogui.press('enter')
            time.sleep(1)

        # PowerPoint: Create a new slide or add text
        elif 'powerpoint' in active_app:
            if 'new slide' in command:
                speak("Adding a new slide in PowerPoint!")
                pyautogui.hotkey('ctrl', 'm')  # Shortcut for new slide
                time.sleep(1)
            elif 'add text' in command:
                text_to_add = command.replace('add text', '').strip()
                speak(f"Adding text '{text_to_add}' to the slide!")
                pyautogui.click(500, 500)  # Click to focus on slide (adjust coordinates if needed)
                pyautogui.write(text_to_add)
      
