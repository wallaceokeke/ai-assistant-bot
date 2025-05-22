import getpass
import sys
import os
import webbrowser
import pyttsx3
import time
import threading
import tkinter as tk
from fuzzywuzzy import fuzz
import psutil
import hashlib
from plyer import notification
from datetime import datetime
import subprocess
import ctypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import wmi
import win32gui
import win32con
import win32api
import win32process
import win32com.client
import pythoncom
import json
import sqlite3
from pathlib import Path
import shutil
import boto3
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import winreg
import socket
import pyautogui
import keyboard
import mouse
import winapps
from fuzzywuzzy import process
import speech_recognition as sr
import re
from threading import Thread, Event
import difflib
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Optional modules
try:
    import pywhatkit
except ImportError:
    print("'pywhatkit' not found. YouTube auto-play disabled.")
    pywhatkit = None

# ---------- NATURAL LANGUAGE PROCESSING ----------
class NaturalLanguageProcessor:
    def __init__(self):
        self.command_patterns = {
            "greeting": ["hi", "hello", "hey", "good morning", "good afternoon", "good evening"],
            "thanks": ["thanks", "thank you", "appreciate it", "much obliged"],
            "farewell": ["bye", "goodbye", "see you", "later"],
            "system": ["system", "computer", "pc", "machine"],
            "browser": ["browser", "chrome", "edge", "firefox", "internet", "web"],
            "app": ["app", "application", "program", "software"],
            "file": ["file", "document", "folder", "directory"],
            "media": ["play", "music", "video", "youtube", "spotify", "netflix", "prime"],
            "control": ["open", "close", "start", "stop", "run", "launch", "go to", "visit"],
            "search": ["search", "find", "look for", "where is"],
            "help": ["help", "what can you do", "show commands", "capabilities"],
            "website": ["youtube", "google", "facebook", "twitter", "instagram", "linkedin", "github"],
            "query": ["what", "when", "where", "who", "why", "how", "search", "find", "look up"],
            "type": ["type", "write", "enter", "input", "paste"],
            "platform": ["google", "chatgpt", "word", "notepad", "browser", "search", "document"],
        }
        
        self.response_templates = {
            "greeting": ["Hi there! How can I help you today?", "Hello! What can I do for you?", "Hey! I'm here to assist you."],
            "thanks": ["You're welcome! Is there anything else you need?", "My pleasure! Let me know if you need anything else.", "Happy to help! What else can I do for you?"],
            "farewell": ["Goodbye! Have a great day!", "See you later! Take care!", "Bye! Come back if you need anything!"],
            "help": [
                "I can help you with many things! Here's what I can do:",
                "- Search the web for information",
                "- Open apps and websites for you",
                "- Type text in documents",
                "- Control your system settings",
                "- Play music and videos",
                "Just let me know what you'd like me to do!"
            ]
        }
    
    def understand_command(self, text):
        """Process natural language command and extract intent"""
        text = text.lower().strip()
        
        # Remove wake word if present
        text = re.sub(r'hey\s+kate\s*', '', text)
        
        # Check for greetings
        if any(greeting in text for greeting in self.command_patterns["greeting"]):
            return {"intent": "greeting", "confidence": 1.0}
        
        # Check for thanks
        if any(thanks in text for thanks in self.command_patterns["thanks"]):
            return {"intent": "thanks", "confidence": 1.0}
        
        # Check for farewell
        if any(farewell in text for farewell in self.command_patterns["farewell"]):
            return {"intent": "farewell", "confidence": 1.0}
        
        # Check for help request
        if any(help_word in text for help_word in self.command_patterns["help"]):
            return {"intent": "help", "confidence": 1.0}
        
        # Extract action and target using simple pattern matching
        action = None
        target = None
        confidence = 0.0
        
        # Find action
        for control_action in self.command_patterns["control"]:
            if control_action in text:
                action = control_action
                confidence = 0.8
                break
        
        # Find target
        for category, patterns in self.command_patterns.items():
            if category not in ["greeting", "thanks", "farewell", "help", "control"]:
                for pattern in patterns:
                    if pattern in text:
                        target = category
                        confidence = (confidence + 0.8) / 2 if confidence > 0 else 0.8
                        break
                if target:
                    break
        
        return {
            "intent": "command",
            "action": action,
            "target": target,
            "confidence": confidence,
            "original_text": text
        }
    
    def get_response(self, intent, confidence=1.0):
        """Generate appropriate response based on intent"""
        if intent in self.response_templates:
            responses = self.response_templates[intent]
            return responses[0] if confidence > 0.8 else responses[-1]
        return "I'm not sure I understand. Could you rephrase that?"

# Initialize NLP
nlp = NaturalLanguageProcessor()

# ---------- VOICE SETTINGS ----------
def setup_voice():
    """Enhanced voice setup with better error handling"""
    try:
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        
        # Find the most natural female voice
        female_voices = [voice for voice in voices if "female" in voice.name.lower()]
        if female_voices:
            # Try to find Zira (Windows default female voice) or any other female voice
            zira_voice = next((voice for voice in female_voices if "zira" in voice.name.lower()), female_voices[0])
            engine.setProperty('voice', zira_voice.id)
        else:
            # If no female voices found, use the first available voice
            engine.setProperty('voice', voices[0].id)
        
        # Adjust voice settings for more natural sound
        engine.setProperty('rate', 150)  # Slightly slower for clarity
        engine.setProperty('volume', 1.0)  # Maximum volume
        
        # Test the voice
        engine.say("Voice test")
        engine.runAndWait()
        
        return engine
    except Exception as e:
        print(f"Error setting up voice: {str(e)}")
        return None

# Initialize voice engine with enhanced settings
engine = setup_voice()

def speak(text):
    """Enhanced speak function with better error handling and feedback"""
    try:
        if engine is None:
            print(f"KATE: {text}")
            return
        
        print(f"KATE: {text}")
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Error in speak function: {str(e)}")
        print(f"KATE: {text}")

# ---------- AUTH ----------
def authenticate():
    correct_password = "wallace123"
    speak("Welcome to KATE, your assistant. Please enter your password.")
    password = getpass.getpass("Enter your password: ")

    if input("👁️ Show typed password? (y/n): ").lower() == 'y':
        print(f"You typed: {password}")

    if password != correct_password:
        speak("Access Denied. This assistant only recognizes Wallace.")
        print(" Access Denied.")
        sys.exit()
    else:
        speak("Access granted. Hello Wallace! How can I help you today?")
        print("Access Granted. Hello Wallace!\n")

# ---------- LISTEN TO VOICE ----------
def listen_command():
    if not sr:
        return input("Voice not available. Type your command: ")

    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        speak("Listening...")
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that.")
        return ""
    except sr.RequestError:
        speak("Voice service error. Please try typing.")
        return ""

# ---------- SYSTEM CONTROL FUNCTIONS ----------
def set_volume(level):
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level / 100, None)
        speak(f"Volume set to {level}%")
    except Exception as e:
        speak("Could not adjust volume")

def set_brightness(level):
    try:
        c = wmi.WMI(namespace='wmi')
        methods = c.WmiMonitorBrightnessMethods()[0]
        methods.WmiSetBrightness(level, 0)
        speak(f"Brightness set to {level}%")
    except Exception as e:
        speak("Could not adjust brightness")

def get_system_info():
    try:
        c = wmi.WMI()
        system_info = c.Win32_ComputerSystem()[0]
        os_info = c.Win32_OperatingSystem()[0]
        cpu_info = c.Win32_Processor()[0]
        
        info = f"""
        System Manufacturer: {system_info.Manufacturer}
        System Model: {system_info.Model}
        Operating System: {os_info.Caption}
        Processor: {cpu_info.Name}
        RAM: {round(int(os_info.TotalVisibleMemorySize) / (1024 * 1024), 2)} GB
        """
        speak(info)
    except Exception as e:
        speak("Could not fetch system information")

def manage_process(action, process_name):
    try:
        if action == "start":
            subprocess.Popen(process_name)
            speak(f"Started {process_name}")
        elif action == "stop":
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower() == process_name.lower():
                    proc.kill()
            speak(f"Stopped {process_name}")
        elif action == "list":
            processes = []
            for proc in psutil.process_iter(['name', 'cpu_percent', 'memory_percent']):
                processes.append(f"{proc.info['name']} - CPU: {proc.info['cpu_percent']}% - Memory: {proc.info['memory_percent']:.1f}%")
            speak("Current running processes:")
            for proc in processes[:10]:  # Limit to top 10
                speak(proc)
    except Exception as e:
        speak(f"Could not {action} process {process_name}")

def network_info():
    try:
        c = wmi.WMI()
        for interface in c.Win32_NetworkAdapterConfiguration(IPEnabled=1):
            speak(f"Network Adapter: {interface.Description}")
            speak(f"IP Address: {interface.IPAddress[0]}")
            speak(f"Subnet Mask: {interface.IPSubnet[0]}")
            speak(f"Default Gateway: {interface.DefaultIPGateway[0]}")
    except Exception as e:
        speak("Could not fetch network information")

# ---------- SMART COMMAND HANDLER ----------
def handle_command(cmd):
    # Process natural language
    understanding = nlp.understand_command(cmd)
    
    if understanding["intent"] == "greeting":
        speak(nlp.get_response("greeting"))
        return
    
    if understanding["intent"] == "thanks":
        speak(nlp.get_response("thanks"))
        return
    
    if understanding["intent"] == "farewell":
        speak(nlp.get_response("farewell"))
        return
    
    if understanding["intent"] == "help":
        for line in nlp.get_response("help"):
            speak(line)
        return
    
    # Handle system commands
    if understanding["intent"] == "command":
        if understanding["confidence"] < 0.5:
            speak("I'm not sure I understand. Could you rephrase that?")
            return
        
        action = understanding["action"]
        target = understanding["target"]
        original_text = understanding["original_text"]
        
        # Query handling
        if any(word in original_text for word in nlp.command_patterns["query"]):
            # Extract the query
            query = re.search(r'(?:what|when|where|who|why|how|search|find|look up)\s+(.*)', original_text)
            if query:
                query_text = query.group(1)
                
                # Determine platform
                platform = "google"  # Default to Google
                if "chatgpt" in original_text:
                    platform = "chatgpt"
                
                # Type the query
                if enhanced_typing.type_in_platform(query_text, platform):
                    speak(f"Searching for {query_text} on {platform}")
                return
        
        # Typing commands
        elif any(word in original_text for word in nlp.command_patterns["type"]):
            # Extract text to type
            text_to_type = re.search(r'(?:type|write|enter|input|paste)\s+(.*)', original_text)
            if text_to_type:
                text = text_to_type.group(1)
                
                # Determine platform
                platform = "current"  # Default to current active window
                if "word" in original_text:
                    platform = "word"
                elif "notepad" in original_text:
                    platform = "notepad"
                elif "google" in original_text:
                    platform = "google"
                elif "chatgpt" in original_text:
                    platform = "chatgpt"
                
                # Type the text
                if enhanced_typing.type_in_platform(text, platform):
                    speak(f"Typing your text in {platform}")
                return
        
        # Website handling
        if any(site in original_text for site in nlp.command_patterns["website"]):
            if "youtube" in original_text:
                if "play" in original_text:
                    # Extract video name if provided
                    video = re.search(r'play\s+(.*?)\s+(?:on|in)\s+youtube', original_text)
                    if video:
                        webbrowser.open(f"https://www.youtube.com/results?search_query={video.group(1)}")
                    else:
                        webbrowser.open("https://www.youtube.com")
                else:
                    webbrowser.open("https://www.youtube.com")
                speak("Opening YouTube for you")
                return
            
            # Handle other common websites
            sites = {
                "google": "https://www.google.com",
                "facebook": "https://www.facebook.com",
                "twitter": "https://www.twitter.com",
                "instagram": "https://www.instagram.com",
                "linkedin": "https://www.linkedin.com",
                "github": "https://www.github.com"
            }
            
            for site, url in sites.items():
                if site in original_text:
                    webbrowser.open(url)
                    speak(f"Opening {site.capitalize()} for you")
                    return
        
        # Media control
        elif target == "media":
            if action == "play":
                media = re.search(r'play\s+(.*)', original_text)
                if media:
                    if "youtube" in original_text:
                        query = media.group(1).replace("on youtube", "").strip()
                        webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
                        speak(f"Searching YouTube for {query}")
                    elif "spotify" in original_text:
                        query = media.group(1).replace("on spotify", "").strip()
                        webbrowser.open(f"https://open.spotify.com/search/{query}")
                        speak(f"Searching Spotify for {query}")
                    else:
                        # Try to play on default media player
                        try:
                            os.startfile(media.group(1))
                            speak(f"Playing {media.group(1)}")
                        except Exception as e:
                            ErrorHandler.handle_error(e, "Media playback")
        
        # System information
        if target == "system":
            if "info" in original_text or "status" in original_text:
                state = system_expert.understand_system_state()
                speak("Here's your system information:")
                for key, value in state["system"].items():
                    speak(f"{key}: {value}")
            elif "process" in original_text or "running" in original_text:
                processes = json.loads(ps_manager.execute_command("Get-Process | Select-Object Name, CPU, WorkingSet | ConvertTo-Json"))
                speak("Running processes:")
                for proc in processes[:5]:  # Show top 5
                    speak(f"- {proc['Name']}: CPU {proc['CPU']}%, Memory {proc['WorkingSet']/1024/1024:.1f}MB")
        
        # Browser control
        elif target == "browser":
            if action == "open":
                browser = next((b for b in ["chrome", "edge", "firefox"] if b in original_text), "default")
                url = re.search(r'open\s+(?:chrome|edge|firefox|browser)\s+(.*)', original_text)
                if url:
                    browser_manager.open_url(url.group(1), browser)
                else:
                    browser_manager.open_url("", browser)
        
        # App control
        elif target == "app":
            if action == "open":
                app_name = re.search(r'open\s+(?:app|application|program)\s+(.*)', original_text)
                if app_name:
                    app_manager.open_app(app_name.group(1))
        
        # File operations
        elif target == "file":
            if action == "open":
                path = re.search(r'open\s+(?:file|document|folder)\s+(.*)', original_text)
                if path:
                    try:
                        os.startfile(path.group(1))
                    except Exception as e:
                        ErrorHandler.handle_error(e, "File opening")
        
        # If no specific handler matched
        else:
            speak("I'm not sure how to help with that. Would you like me to list what I can do?")
    
    # If no intent was recognized
    else:
        speak("I'm not sure I understand. Could you rephrase that?")

# ---------- BACKGROUND ENHANCED MONITOR ----------
IDLE_APPS = ["chrome.exe", "discord.exe", "windscribe.exe"]
IDLE_LIMIT = 1800
last_cpu_usage = {}

def show_countdown_popup(app_name, pid):
    def kill_app():
        try:
            os.kill(pid, 9)
            speak(f"{app_name} has been paused after being idle. You can resume it later.")
        except Exception as e:
            print(f"Error closing {app_name}: {e}")

    def cancel():
        print(f"❗ {app_name} will stay open.")
        window.destroy()

    countdown = 10
    window = tk.Tk()
    window.title("Idle App Detected")
    window.geometry("400x150")
    tk.Label(window, text=f"{app_name} has been idle.\nClosing in 10 seconds...", font=("Arial", 12)).pack(pady=10)
    cancel_button = tk.Button(window, text="Cancel", command=cancel, bg="red", fg="white")
    cancel_button.pack()

    def tick():
        nonlocal countdown
        if countdown == 0:
            window.destroy()
            kill_app()
        else:
            window.children['!label'].config(text=f"{app_name} has been idle.\nClosing in {countdown} seconds...")
            countdown -= 1
            window.after(1000, tick)

    tick()
    window.mainloop()

def enhanced_idle_monitor():
    while True:
        for proc in psutil.process_iter(['pid', 'name', 'cpu_times']):
            try:
                name = proc.info['name'].lower()
                pid = proc.info['pid']
                cpu = proc.info['cpu_times'].user

                if name in IDLE_APPS:
                    if name not in last_cpu_usage:
                        last_cpu_usage[name] = (cpu, time.time())
                    else:
                        old_cpu, last_time = last_cpu_usage[name]
                        if cpu == old_cpu and time.time() - last_time >= IDLE_LIMIT:
                            threading.Thread(target=show_countdown_popup, args=(name, pid), daemon=True).start()
                            last_cpu_usage[name] = (cpu, time.time() + 60)
                        elif cpu != old_cpu:
                            last_cpu_usage[name] = (cpu, time.time())
            except:
                continue
        time.sleep(60)

# ---------- DUPLICATE FILE CLEANER ----------
def clean_duplicates():
    scanned = {}
    docs_path = os.path.expanduser("~\\Documents")
    for root, _, files in os.walk(docs_path):
        for file in files:
            try:
                path = os.path.join(root, file)
                with open(path, 'rb') as f:
                    file_hash = hashlib.md5(f.read()).hexdigest()
                if file_hash in scanned:
                    os.remove(path)
                    speak(f"Deleted duplicate: {file}")
                else:
                    scanned[file_hash] = path
            except:
                continue

# ---------- USER PATTERN TRACKING ----------
class UserPatternTracker:
    def __init__(self):
        self.db_path = "user_patterns.db"
        self.init_db()
        self.current_session = {
            'start_time': datetime.now(),
            'active_apps': {},
            'documents_opened': set()
        }

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS app_usage
                    (app_name TEXT, duration INTEGER, date TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS document_usage
                    (doc_path TEXT, last_accessed TEXT, frequency INTEGER)''')
        c.execute('''CREATE TABLE IF NOT EXISTS daily_patterns
                    (date TEXT, most_active_hours TEXT, favorite_apps TEXT)''')
        conn.commit()
        conn.close()

    def track_app_usage(self, app_name):
        if app_name not in self.current_session['active_apps']:
            self.current_session['active_apps'][app_name] = datetime.now()
        else:
            duration = (datetime.now() - self.current_session['active_apps'][app_name]).total_seconds()
            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()
            c.execute("INSERT INTO app_usage VALUES (?, ?, ?)",
                     (app_name, duration, datetime.now().strftime("%Y-%m-%d")))
            conn.commit()
            conn.close()

    def track_document_usage(self, doc_path):
        self.current_session['documents_opened'].add(doc_path)
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("INSERT OR REPLACE INTO document_usage VALUES (?, ?, COALESCE((SELECT frequency FROM document_usage WHERE doc_path = ?), 0) + 1)",
                 (doc_path, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), doc_path))
        conn.commit()
        conn.close()

    def get_suggestions(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # Get most used apps
        c.execute("SELECT app_name, SUM(duration) as total FROM app_usage GROUP BY app_name ORDER BY total DESC LIMIT 3")
        favorite_apps = [row[0] for row in c.fetchall()]
        
        # Get recently accessed documents
        c.execute("SELECT doc_path FROM document_usage ORDER BY last_accessed DESC LIMIT 5")
        recent_docs = [row[0] for row in c.fetchall()]
        
        conn.close()
        
        return {
            'favorite_apps': favorite_apps,
            'recent_docs': recent_docs
        }

# ---------- IDLE APP MONITOR ----------
class IdleAppMonitor:
    def __init__(self):
        self.idle_threshold = 1800  # 30 minutes
        self.last_activity = {}
        self.browser_tabs = {}

    def check_idle_apps(self):
        current_time = time.time()
        for proc in psutil.process_iter(['pid', 'name', 'create_time']):
            try:
                proc_info = proc.info
                name = proc_info['name'].lower()
                pid = proc_info['pid']
                
                if name not in self.last_activity:
                    self.last_activity[name] = current_time
                else:
                    idle_time = current_time - self.last_activity[name]
                    if idle_time > self.idle_threshold:
                        self.handle_idle_app(name, pid)
            except:
                continue

    def handle_idle_app(self, app_name, pid):
        speak(f"I noticed {app_name} has been idle for a while. Would you like me to close it to save resources?")
        response = input("Close idle app? (y/n): ").lower()
        if response == 'y':
            try:
                os.kill(pid, 9)
                speak(f"Closed {app_name}")
            except:
                speak(f"Could not close {app_name}")

# ---------- DOCUMENT MONITOR ----------
class DocumentMonitor(FileSystemEventHandler):
    def __init__(self):
        self.cloud_storage = None
        self.setup_cloud_storage()

    def setup_cloud_storage(self):
        try:
            # Initialize AWS S3 client (you'll need to configure credentials)
            self.cloud_storage = boto3.client('s3')
        except:
            speak("Cloud storage not configured. Please set up your cloud storage credentials.")

    def on_modified(self, event):
        if not event.is_directory:
            self.check_document_age(event.src_path)

    def check_document_age(self, file_path):
        try:
            last_modified = os.path.getmtime(file_path)
            age_days = (time.time() - last_modified) / (24 * 3600)
            
            if age_days > 90:  # 90 days old
                speak(f"I found an old document: {os.path.basename(file_path)}. Would you like to move it to cloud storage?")
                response = input("Move to cloud? (y/n): ").lower()
                if response == 'y':
                    self.move_to_cloud(file_path)
        except:
            pass

    def move_to_cloud(self, file_path):
        if self.cloud_storage:
            try:
                bucket_name = "your-bucket-name"  # Replace with your bucket name
                key = os.path.basename(file_path)
                self.cloud_storage.upload_file(file_path, bucket_name, key)
                speak(f"Moved {key} to cloud storage")
            except:
                speak("Failed to move file to cloud storage")

# ---------- STARTUP GREETING ----------
def startup_greeting():
    hour = datetime.now().hour
    if 5 <= hour < 12:
        greeting = "Good morning"
    elif 12 <= hour < 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"
    
    speak(f"{greeting} Wallace! Welcome back. Having a great day?")
    time.sleep(1)
    
    # Get suggestions based on patterns
    suggestions = pattern_tracker.get_suggestions()
    if suggestions['favorite_apps']:
        speak("Based on your patterns, you might want to work with:")
        for app in suggestions['favorite_apps']:
            speak(app)
    
    speak("What would you like to do today?")

# Initialize trackers
pattern_tracker = UserPatternTracker()
idle_monitor = IdleAppMonitor()
doc_monitor = DocumentMonitor()

# Start document monitoring
observer = Observer()
observer.schedule(doc_monitor, path=os.path.expanduser("~\\Documents"), recursive=True)
observer.start()

# ---------- APP INTEGRATION ----------
class AppIntegration:
    def __init__(self):
        self.app_handles = {}
        self.setup_app_integration()

    def setup_app_integration(self):
        # Register for common applications
        self.register_app("vscode", "Code.exe")
        self.register_app("cursor", "Cursor.exe")
        self.register_app("chrome", "chrome.exe")
        self.register_app("notepad", "notepad.exe")

    def register_app(self, app_name, process_name):
        self.app_handles[app_name] = process_name

    def control_app(self, app_name, action):
        if app_name in self.app_handles:
            process_name = self.app_handles[app_name]
            if action == "start":
                subprocess.Popen(process_name)
                speak(f"Starting {app_name}")
            elif action == "stop":
                for proc in psutil.process_iter(['name']):
                    if proc.info['name'].lower() == process_name.lower():
                        proc.kill()
                speak(f"Stopping {app_name}")
            elif action == "focus":
                try:
                    hwnd = win32gui.FindWindow(None, app_name)
                    if hwnd:
                        win32gui.SetForegroundWindow(hwnd)
                        speak(f"Focusing on {app_name}")
                except:
                    speak(f"Could not focus on {app_name}")

# Initialize app integration
app_integrator = AppIntegration()

# ---------- STARTUP INTEGRATION ----------
def setup_startup():
    try:
        # Create a startup script
        startup_script = """
@echo off
powershell -ExecutionPolicy Bypass -File "C:\AI_Assistant\launcher.ps1"
"""
        with open("startup.bat", "w") as f:
            f.write(startup_script)

        # Add to Windows startup
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(
            key,
            "KATEAssistant",
            0,
            winreg.REG_SZ,
            os.path.abspath("startup.bat")
        )
        winreg.CloseKey(key)
        speak("KATE Assistant has been added to startup")
    except Exception as e:
        speak(f"Could not add to startup: {str(e)}")

# ---------- TAB MANAGEMENT ----------
class TabManager:
    def __init__(self):
        self.idle_tabs = {}
        self.idle_threshold = 1800  # 30 minutes
        self.countdown = 30  # seconds

    def monitor_tabs(self):
        while True:
            try:
                for proc in psutil.process_iter(['pid', 'name', 'create_time']):
                    try:
                        if proc.info['name'].lower() in ['chrome.exe', 'msedge.exe', 'firefox.exe']:
                            self.check_tab_activity(proc)
                    except:
                        continue
                time.sleep(60)  # Check every minute
            except:
                continue

    def check_tab_activity(self, proc):
        pid = proc.info['pid']
        if pid not in self.idle_tabs:
            self.idle_tabs[pid] = time.time()
        else:
            idle_time = time.time() - self.idle_tabs[pid]
            if idle_time > self.idle_threshold:
                self.handle_idle_tabs(proc)

    def handle_idle_tabs(self, proc):
        speak(f"I noticed {proc.info['name']} has been idle for a while. Would you like me to close it to save resources?")
        speak(f"You have {self.countdown} seconds to respond.")
        
        def countdown():
            for i in range(self.countdown, 0, -1):
                if i % 10 == 0:  # Announce every 10 seconds
                    speak(f"{i} seconds remaining")
                time.sleep(1)
            return True

        def wait_for_response():
            response = input("Close idle browser? (y/n): ").lower()
            return response == 'y'

        # Run countdown and response check in parallel
        countdown_thread = threading.Thread(target=countdown)
        response_thread = threading.Thread(target=wait_for_response)
        
        countdown_thread.start()
        response_thread.start()
        
        countdown_thread.join()
        response_thread.join()
        
        if response_thread.is_alive():
            speak("Time's up! Closing idle browser.")
            try:
                proc.kill()
                speak(f"Closed {proc.info['name']}")
            except:
                speak(f"Could not close {proc.info['name']}")

# ---------- ENHANCED APP CONTROL ----------
class EnhancedAppControl:
    def __init__(self):
        self.app_paths = {
            'cursor': r"C:\Users\{}\AppData\Local\Programs\Cursor\Cursor.exe",
            'vscode': r"C:\Users\{}\AppData\Local\Programs\Microsoft VS Code\Code.exe",
            'spotify': r"C:\Users\{}\AppData\Roaming\Spotify\Spotify.exe",
            'pictures': r"C:\Users\{}\Pictures",
            'downloads': r"C:\Users\{}\Downloads",
            'documents': r"C:\Users\{}\Documents",
            'assistant': r"C:\AI_Assistant"
        }
        self.username = os.getenv('USERNAME')

    def open_app(self, app_name):
        if app_name in self.app_paths:
            path = self.app_paths[app_name].format(self.username)
            try:
                if os.path.isdir(path):
                    os.startfile(path)
                else:
                    subprocess.Popen(path)
                speak(f"Opening {app_name}")
            except Exception as e:
                speak(f"Could not open {app_name}: {str(e)}")
        else:
            speak(f"Sorry, I don't know how to open {app_name}")

    def play_media(self, platform, query):
        if platform == "youtube":
            if pywhatkit:
                speak(f"Playing {query} on YouTube")
                pywhatkit.playonyt(query)
            else:
                query = '+'.join(query.split())
                webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
        elif platform == "spotify":
            query = '%20'.join(query.split())
            webbrowser.open(f"https://open.spotify.com/search/{query}")

# Initialize managers
tab_manager = TabManager()
app_controller = EnhancedAppControl()

# ---------- SYSTEM AUTOMATION ----------
class SystemAutomation:
    def __init__(self):
        self.window_handles = {}
        self.screen_width, self.screen_height = pyautogui.size()
        self.mouse_speed = 0.5  # Slower mouse movement for precision
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1

    def move_mouse(self, x, y):
        """Move mouse to coordinates with smooth animation"""
        current_x, current_y = pyautogui.position()
        pyautogui.moveTo(x, y, duration=self.mouse_speed)

    def click(self, x=None, y=None):
        """Click at specified coordinates or current position"""
        if x is not None and y is not None:
            self.move_mouse(x, y)
        pyautogui.click()

    def double_click(self, x=None, y=None):
        """Double click at specified coordinates or current position"""
        if x is not None and y is not None:
            self.move_mouse(x, y)
        pyautogui.doubleClick()

    def right_click(self, x=None, y=None):
        """Right click at specified coordinates or current position"""
        if x is not None and y is not None:
            self.move_mouse(x, y)
        pyautogui.rightClick()

    def scroll(self, amount):
        """Scroll up (positive) or down (negative)"""
        pyautogui.scroll(amount)

    def type_text(self, text):
        """Type text with natural timing"""
        pyautogui.write(text, interval=0.1)

    def press_key(self, key):
        """Press a single key"""
        pyautogui.press(key)

    def hotkey(self, *keys):
        """Press multiple keys in sequence"""
        pyautogui.hotkey(*keys)

    def find_window(self, title):
        """Find window by title and return its handle"""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    self.window_handles[title] = hwnd
        win32gui.EnumWindows(callback, None)
        return self.window_handles.get(title)

    def focus_window(self, title):
        """Bring window to front and focus"""
        hwnd = self.find_window(title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            return True
        return False

    def move_window(self, title, x, y, width, height):
        """Move and resize window"""
        hwnd = self.find_window(title)
        if hwnd:
            win32gui.MoveWindow(hwnd, x, y, width, height, True)
            return True
        return False

    def maximize_window(self, title):
        """Maximize window"""
        hwnd = self.find_window(title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            return True
        return False

    def minimize_window(self, title):
        """Minimize window"""
        hwnd = self.find_window(title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            return True
        return False

    def close_window(self, title):
        """Close window"""
        hwnd = self.find_window(title)
        if hwnd:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return True
        return False

    def get_active_window(self):
        """Get title of currently active window"""
        return win32gui.GetWindowText(win32gui.GetForegroundWindow())

    def take_screenshot(self, region=None):
        """Take screenshot of specified region or full screen"""
        if region:
            return pyautogui.screenshot(region=region)
        return pyautogui.screenshot()

# Initialize system automation
system_automation = SystemAutomation()

# ---------- ERROR HANDLING ----------
class ErrorHandler:
    @staticmethod
    def handle_error(error, context=""):
        """Handle errors gracefully with user-friendly messages"""
        error_messages = {
            "app_not_found": "I couldn't find that application. Would you like me to search for something similar?",
            "email_error": "I'm having trouble with email. Please make sure Outlook is set up correctly.",
            "command_error": "I'm not sure about that command. Let me suggest some alternatives.",
            "system_error": "I encountered a system error. Let me try a different approach.",
            "voice_error": "I'm having trouble with voice recognition. Would you like to type instead?",
            "file_error": "I couldn't access that file. Please check if it exists and try again.",
            "permission_error": "I don't have permission to do that. Please try running as administrator.",
            "connection_error": "I'm having trouble connecting. Please check your internet connection.",
            "general_error": "Something went wrong. Let me try to fix it and continue."
        }
        
        # Get appropriate error message
        if isinstance(error, FileNotFoundError):
            message = error_messages["file_error"]
        elif isinstance(error, PermissionError):
            message = error_messages["permission_error"]
        elif "outlook" in str(error).lower():
            message = error_messages["email_error"]
        elif "app" in str(error).lower():
            message = error_messages["app_not_found"]
        else:
            message = error_messages["general_error"]
        
        speak(message)
        return message

# ---------- ENHANCED APP INTEGRATION ----------
class EnhancedAppManager:
    def __init__(self):
        self.installed_apps = {}
        self.user_patterns = {}
        self.error_handler = ErrorHandler()
        try:
            self.load_installed_apps()
            self.load_user_patterns()
        except Exception as e:
            self.error_handler.handle_error(e, "AppManager initialization")

    def load_installed_apps(self):
        """Load all installed applications from the system"""
        try:
            for app in winapps.list_installed():
                if app.name:
                    self.installed_apps[app.name.lower()] = app.install_location
        except Exception as e:
            # Fallback to registry
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                for i in range(0, winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, subkey) as app_key:
                            name = winreg.QueryValueEx(app_key, "DisplayName")[0]
                            path = winreg.QueryValueEx(app_key, "InstallLocation")[0]
                            if name and path:
                                self.installed_apps[name.lower()] = path
                    except:
                        continue
            except Exception as e:
                self.error_handler.handle_error(e, "Registry app loading")
                # Add common apps as fallback
                self.installed_apps.update({
                    "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    "notepad": r"C:\Windows\System32\notepad.exe",
                    "calculator": r"C:\Windows\System32\calc.exe"
                })

    def load_user_patterns(self):
        """Load user usage patterns from file"""
        try:
            with open("user_patterns.json", "r") as f:
                self.user_patterns = json.load(f)
        except Exception as e:
            self.error_handler.handle_error(e, "Pattern loading")
            self.user_patterns = {"frequent_apps": {}, "common_commands": {}}

    def save_user_patterns(self):
        """Save user usage patterns to file"""
        with open("user_patterns.json", "w") as f:
            json.dump(self.user_patterns, f)

    def update_patterns(self, app_name, command):
        """Update usage patterns"""
        if app_name not in self.user_patterns["frequent_apps"]:
            self.user_patterns["frequent_apps"][app_name] = 0
        self.user_patterns["frequent_apps"][app_name] += 1
        
        if command not in self.user_patterns["common_commands"]:
            self.user_patterns["common_commands"][command] = 0
        self.user_patterns["common_commands"][command] += 1
        
        self.save_user_patterns()

    def get_suggestions(self, command):
        """Get suggestions based on user patterns"""
        suggestions = []
        # Get frequent apps
        frequent_apps = sorted(self.user_patterns["frequent_apps"].items(), 
                             key=lambda x: x[1], reverse=True)[:3]
        suggestions.extend([f"open {app}" for app, _ in frequent_apps])
        
        # Get similar commands
        similar_commands = process.extract(command, 
                                         self.user_patterns["common_commands"].keys(),
                                         limit=3)
        suggestions.extend([cmd for cmd, _ in similar_commands])
        
        return suggestions

    def open_app(self, app_name):
        """Open any installed application with error handling"""
        app_name = app_name.lower()
        try:
            # First try exact match
            if app_name in self.installed_apps:
                try:
                    subprocess.Popen(self.installed_apps[app_name])
                    self.update_patterns(app_name, f"open {app_name}")
                    return True
                except Exception as e:
                    self.error_handler.handle_error(e, f"Opening app {app_name}")
            
            # Try fuzzy matching
            matches = process.extract(app_name, self.installed_apps.keys(), limit=3)
            if matches:
                speak(f"I couldn't find {app_name}. Did you mean:")
                for match, score in matches:
                    if score > 70:
                        speak(f"- {match}")
                return True
            
            return False
        except Exception as e:
            self.error_handler.handle_error(e, "App opening")
            return False

# ---------- EMAIL INTEGRATION ----------
class EmailManager:
    def __init__(self):
        self.outlook = None
        self.error_handler = ErrorHandler()
        self.setup_outlook()

    def setup_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            self.error_handler.handle_error(e, "Outlook setup")

    def send_email(self, to, subject, body):
        try:
            if not self.outlook:
                raise Exception("Outlook not connected")
            mail = self.outlook.CreateItem(0)
            mail.To = to
            mail.Subject = subject
            mail.Body = body
            mail.Send()
            return True
        except Exception as e:
            self.error_handler.handle_error(e, "Email sending")
            return False

    def get_saved_accounts(self):
        try:
            accounts = []
            for account in self.outlook.Session.Accounts:
                accounts.append(account.SmtpAddress)
            return accounts
        except:
            return []

# Initialize enhanced managers
app_manager = EnhancedAppManager()
email_manager = EmailManager()

# ---------- WAKE WORD DETECTION ----------
class InteractionManager:
    def __init__(self):
        self.last_interaction = time.time()
        self.is_awake = False
        self.silence_threshold = 30  # 30 seconds of silence before asking
        self.sleep_threshold = 300  # 5 minutes of silence before checking
        self.auto_sleep_time = 30  # 30 seconds before auto-sleep
        self.has_greeted = False
        self.has_listened = False
    
    def update_interaction(self):
        self.last_interaction = time.time()
    
    def check_silence(self):
        current_time = time.time()
        silence_duration = current_time - self.last_interaction
        
        if not self.is_awake:
            return False
        
        if silence_duration > self.sleep_threshold:
            speak("Are you still there?")
            time.sleep(self.auto_sleep_time)
            if current_time - self.last_interaction > self.auto_sleep_time:
                speak("Going to sleep. Say 'Hey Kate' when you need me.")
                self.is_awake = False
                self.has_greeted = False
                self.has_listened = False
                return True
            self.update_interaction()
        
        elif silence_duration > self.silence_threshold and not self.has_listened:
            speak("I'm here if you need me.")
            self.has_listened = True
        
        return False

class WakeWordDetector:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = None
        self.is_listening = False
        self.stop_event = Event()
        self.wake_word = "hey kate"
        self.sleep_word = "sleep"
        self.interaction_manager = InteractionManager()
        
        try:
            self.microphone = sr.Microphone()
            with self.microphone as source:
                print("\n🎤 Adjusting for ambient noise...")
                self.recognizer.adjust_for_ambient_noise(source, duration=2)
                print("✅ Microphone calibrated!")
        except Exception as e:
            print(f"\n❌ Error initializing microphone: {str(e)}")
            print("Please check your microphone connection and try again.")
            raise
    
    def start_listening(self):
        """Start continuous listening for wake word"""
        if not self.microphone:
            print("\n❌ Microphone not available. Please check your audio devices.")
            return
            
        self.is_listening = True
        self.stop_event.clear()
        self.interaction_manager.is_awake = True
        
        def listen_loop():
            while not self.stop_event.is_set():
                try:
                    with self.microphone as source:
                        print("\n🎤 Listening... (Say 'Hey Kate' to activate)", end='\r')
                        audio = self.recognizer.listen(source, timeout=1, phrase_time_limit=3)
                        try:
                            text = self.recognizer.recognize_google(audio).lower()
                            self.interaction_manager.update_interaction()
                            
                            if self.wake_word in text:
                                print("\n👂 Heard 'Hey Kate'!")  # Clear feedback when wake word detected
                                if not self.interaction_manager.has_greeted:
                                    speak("Hello! I'm KATE, your personal assistant. How can I help you today?")
                                    self.interaction_manager.has_greeted = True
                                else:
                                    speak("Yes, I'm listening. What can I do for you?")
                                handle_command(text.replace(self.wake_word, "").strip())
                            elif self.sleep_word in text:
                                print("\n😴 Going to sleep...")  # Clear feedback when going to sleep
                                speak("Going to sleep. Say 'Hey Kate' when you need me.")
                                self.interaction_manager.is_awake = False
                                self.interaction_manager.has_greeted = False
                                self.stop_event.set()
                        except sr.UnknownValueError:
                            if self.interaction_manager.check_silence():
                                self.stop_event.set()
                            continue
                        except sr.RequestError as e:
                            print(f"\n⚠️ Voice service error: {str(e)}")
                            print("Retrying in 5 seconds...")
                            time.sleep(5)
                            continue
                except Exception as e:
                    print(f"\n⚠️ Error: {str(e)}")
                    print("Retrying in 5 seconds...")
                    time.sleep(5)
                    continue
        
        Thread(target=listen_loop, daemon=True).start()
    
    def stop_listening(self):
        """Stop continuous listening"""
        self.stop_event.set()
        self.is_listening = False
        self.interaction_manager.is_awake = False
        print("\n🔇 Listening stopped")

# ---------- ENHANCED POWERSHELL INTEGRATION ----------
class PowerShellManager:
    def __init__(self):
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.processes = {}
    
    def execute_command(self, command):
        """Execute PowerShell command and return output"""
        try:
            process = subprocess.Popen(
                ["powershell", "-Command", command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            output, error = process.communicate()
            return output.strip() if not error else f"Error: {error.strip()}"
        except Exception as e:
            return f"Error executing command: {str(e)}"
    
    def get_installed_apps(self):
        """Get all installed applications using PowerShell"""
        command = """
        Get-ItemProperty HKLM:\\Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | 
        Select-Object DisplayName, InstallLocation | 
        Where-Object { $_.DisplayName -ne $null } | 
        ConvertTo-Json
        """
        return self.execute_command(command)
    
    def get_running_processes(self):
        """Get detailed process information"""
        command = "Get-Process | Select-Object Name, Id, CPU, WorkingSet | ConvertTo-Json"
        return self.execute_command(command)
    
    def open_browser(self, url, browser="default"):
        """Open URL in specified browser"""
        browsers = {
            "chrome": "chrome.exe",
            "edge": "msedge.exe",
            "firefox": "firefox.exe",
            "default": "start"
        }
        if browser in browsers:
            self.shell.Run(f"{browsers[browser]} {url}")
    
    def open_app_store(self, store="playstore"):
        """Open app store"""
        stores = {
            "playstore": "https://play.google.com/store",
            "microsoft": "ms-windows-store://",
            "apple": "https://apps.apple.com"
        }
        if store in stores:
            self.open_browser(stores[store])
    
    def open_whatsapp(self):
        """Open WhatsApp Web"""
        self.open_browser("https://web.whatsapp.com")

# ---------- ENHANCED BROWSER INTEGRATION ----------
class BrowserManager:
    def __init__(self):
        self.ps_manager = PowerShellManager()
        self.browsers = {
            "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            "firefox": r"C:\Program Files\Mozilla Firefox\firefox.exe"
        }
    
    def open_url(self, url, browser="default"):
        """Open URL in specified browser"""
        if browser == "default":
            webbrowser.open(url)
        elif browser in self.browsers:
            subprocess.Popen([self.browsers[browser], url])
    
    def open_whatsapp(self):
        """Open WhatsApp Web"""
        self.open_url("https://web.whatsapp.com")
    
    def open_app_store(self, store="playstore"):
        """Open app store"""
        stores = {
            "playstore": "https://play.google.com/store",
            "microsoft": "ms-windows-store://",
            "apple": "https://apps.apple.com"
        }
        if store in stores:
            self.open_url(stores[store])

# Initialize managers
wake_detector = WakeWordDetector()
ps_manager = PowerShellManager()
browser_manager = BrowserManager()

# ---------- ENHANCED SYSTEM INTEGRATION ----------
class SystemExpert:
    def __init__(self):
        self.ps_manager = PowerShellManager()
        self.nlp = NaturalLanguageProcessor()
        self.system_commands = {
            "process": {
                "list": "Get-Process | Select-Object Name, Id, CPU, WorkingSet | ConvertTo-Json",
                "start": "Start-Process -FilePath '{path}'",
                "stop": "Stop-Process -Name '{name}' -Force"
            },
            "service": {
                "list": "Get-Service | Select-Object Name, Status, StartType | ConvertTo-Json",
                "start": "Start-Service -Name '{name}'",
                "stop": "Stop-Service -Name '{name}' -Force"
            },
            "network": {
                "info": "Get-NetIPConfiguration | Select-Object InterfaceAlias, IPv4Address, IPv4DefaultGateway | ConvertTo-Json",
                "adapters": "Get-NetAdapter | Select-Object Name, InterfaceDescription, Status | ConvertTo-Json"
            },
            "system": {
                "info": "Get-ComputerInfo | Select-Object WindowsProductName, WindowsVersion, OsHardwareAbstractionLayer | ConvertTo-Json",
                "uptime": "Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object LastBootUpTime | ConvertTo-Json"
            }
        }
    
    def execute_system_command(self, category, action, params=None):
        """Execute system command with proper parameters"""
        if category in self.system_commands and action in self.system_commands[category]:
            command = self.system_commands[category][action]
            if params:
                command = command.format(**params)
            return self.ps_manager.execute_command(command)
        return "Command not found"
    
    def understand_system_state(self):
        """Get comprehensive system state"""
        state = {
            "processes": json.loads(self.execute_system_command("process", "list")),
            "services": json.loads(self.execute_system_command("service", "list")),
            "network": json.loads(self.execute_system_command("network", "info")),
            "system": json.loads(self.execute_system_command("system", "info"))
        }
        return state

# Initialize enhanced components
nlp = NaturalLanguageProcessor()
system_expert = SystemExpert()

# ---------- ENHANCED TYPING ----------
class EnhancedTyping:
    def __init__(self):
        self.driver = None
        self.current_platform = None
        self.setup_selenium()
    
    def setup_selenium(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--start-maximized')
            self.driver = webdriver.Chrome(options=options)
        except Exception as e:
            print(f"Could not initialize Chrome driver: {e}")
            self.driver = None
    
    def type_in_platform(self, text, platform):
        """Type text in specified platform"""
        try:
            if platform == "google":
                if not self.driver:
                    webbrowser.open("https://www.google.com")
                    time.sleep(2)  # Wait for page to load
                else:
                    self.driver.get("https://www.google.com")
                    search_box = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.NAME, "q"))
                    )
                    search_box.clear()
                    search_box.send_keys(text)
                    search_box.send_keys(Keys.RETURN)
            
            elif platform == "chatgpt":
                if not self.driver:
                    webbrowser.open("https://chat.openai.com")
                    time.sleep(2)
                else:
                    self.driver.get("https://chat.openai.com")
                    # Wait for chat input to be available
                    chat_input = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "textarea"))
                    )
                    chat_input.clear()
                    chat_input.send_keys(text)
                    chat_input.send_keys(Keys.RETURN)
            
            elif platform in ["word", "notepad"]:
                # Use pyautogui to type in any application
                pyperclip.copy(text)  # Copy text to clipboard
                time.sleep(0.5)  # Small delay
                keyboard.press_and_release('ctrl+v')  # Paste text
            
            else:
                # Generic typing for any platform
                pyperclip.copy(text)
                time.sleep(0.5)
                keyboard.press_and_release('ctrl+v')
            
            return True
        except Exception as e:
            ErrorHandler.handle_error(e, f"Typing in {platform}")
            return False

# Initialize enhanced typing
enhanced_typing = EnhancedTyping()

# ---------- ENHANCED FILE MANAGEMENT ----------
class FileManager:
    def __init__(self):
        self.user_home = os.path.expanduser("~")
        self.common_paths = {
            "desktop": os.path.join(self.user_home, "Desktop"),
            "downloads": os.path.join(self.user_home, "Downloads"),
            "documents": os.path.join(self.user_home, "Documents"),
            "pictures": os.path.join(self.user_home, "Pictures"),
            "music": os.path.join(self.user_home, "Music"),
            "videos": os.path.join(self.user_home, "Videos")
        }
        self.recent_actions = []
        self.setup_file_operations()

    def setup_file_operations(self):
        """Initialize file operation handlers"""
        self.operations = {
            "create": self.create_item,
            "delete": self.delete_item,
            "move": self.move_item,
            "copy": self.copy_item,
            "rename": self.rename_item,
            "share": self.share_item,
            "download": self.download_file,
            "empty": self.empty_folder,
            "shortcut": self.create_shortcut
        }

    def create_item(self, name, item_type="folder", location="downloads"):
        """Create a folder or file"""
        try:
            if location in self.common_paths:
                path = os.path.join(self.common_paths[location], name)
                if item_type == "folder":
                    os.makedirs(path, exist_ok=True)
                    speak(f"Created folder {name} in {location}")
                else:
                    with open(path, "w") as f:
                        f.write("")
                    speak(f"Created file {name} in {location}")
                self.recent_actions.append(f"Created {item_type} {name} in {location}")
                return True
            return False
        except Exception as e:
            speak(f"Could not create {item_type}: {str(e)}")
            return False

    def delete_item(self, name, location="downloads"):
        """Delete a file or folder"""
        try:
            if location in self.common_paths:
                path = os.path.join(self.common_paths[location], name)
                if os.path.exists(path):
                    if os.path.isfile(path):
                        os.remove(path)
                    else:
                        shutil.rmtree(path)
                    speak(f"Deleted {name} from {location}")
                    self.recent_actions.append(f"Deleted {name} from {location}")
                    return True
            return False
        except Exception as e:
            speak(f"Could not delete item: {str(e)}")
            return False

    def move_item(self, name, from_location, to_location):
        """Move a file or folder"""
        try:
            if from_location in self.common_paths and to_location in self.common_paths:
                from_path = os.path.join(self.common_paths[from_location], name)
                to_path = os.path.join(self.common_paths[to_location], name)
                if os.path.exists(from_path):
                    shutil.move(from_path, to_path)
                    speak(f"Moved {name} from {from_location} to {to_location}")
                    self.recent_actions.append(f"Moved {name} from {from_location} to {to_location}")
                    return True
            return False
        except Exception as e:
            speak(f"Could not move item: {str(e)}")
            return False

    def copy_item(self, name, from_location, to_location):
        """Copy a file or folder"""
        try:
            if from_location in self.common_paths and to_location in self.common_paths:
                from_path = os.path.join(self.common_paths[from_location], name)
                to_path = os.path.join(self.common_paths[to_location], name)
                if os.path.exists(from_path):
                    if os.path.isfile(from_path):
                        shutil.copy2(from_path, to_path)
                    else:
                        shutil.copytree(from_path, to_path)
                    speak(f"Copied {name} from {from_location} to {to_location}")
                    self.recent_actions.append(f"Copied {name} from {from_location} to {to_location}")
                    return True
            return False
        except Exception as e:
            speak(f"Could not copy item: {str(e)}")
            return False

    def rename_item(self, old_name, new_name, location="downloads"):
        """Rename a file or folder"""
        try:
            if location in self.common_paths:
                old_path = os.path.join(self.common_paths[location], old_name)
                new_path = os.path.join(self.common_paths[location], new_name)
                if os.path.exists(old_path):
                    os.rename(old_path, new_path)
                    speak(f"Renamed {old_name} to {new_name}")
                    self.recent_actions.append(f"Renamed {old_name} to {new_name}")
                    return True
            return False
        except Exception as e:
            speak(f"Could not rename item: {str(e)}")
            return False

    def empty_folder(self, location="downloads"):
        """Empty a folder"""
        try:
            if location in self.common_paths:
                path = self.common_paths[location]
                for item in os.listdir(path):
                    item_path = os.path.join(path, item)
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                    else:
                        shutil.rmtree(item_path)
                speak(f"Emptied {location} folder")
                self.recent_actions.append(f"Emptied {location} folder")
                return True
            return False
        except Exception as e:
            speak(f"Could not empty folder: {str(e)}")
            return False

    def create_shortcut(self, target_name, shortcut_name, location="desktop"):
        """Create a desktop shortcut"""
        try:
            if location in self.common_paths:
                target_path = os.path.join(self.common_paths["downloads"], target_name)
                shortcut_path = os.path.join(self.common_paths[location], f"{shortcut_name}.lnk")
                
                pythoncom.CoInitialize()
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target_path
                shortcut.save()
                
                speak(f"Created shortcut {shortcut_name} on {location}")
                self.recent_actions.append(f"Created shortcut {shortcut_name} on {location}")
                return True
            return False
        except Exception as e:
            speak(f"Could not create shortcut: {str(e)}")
            return False

    def handle_command(self, command):
        """Process file management commands"""
        command = command.lower()
        
        # Extract operation and parameters
        for operation, handler in self.operations.items():
            if operation in command:
                # Extract parameters based on operation
                if operation == "create":
                    if "folder" in command:
                        name = command.replace("create folder", "").strip()
                        location = "downloads"
                        if " in " in name:
                            name, location = name.split(" in ")
                            name = name.strip()
                            location = location.strip()
                        return handler(name, "folder", location)
                    else:
                        name = command.replace("create file", "").strip()
                        location = "downloads"
                        if " in " in name:
                            name, location = name.split(" in ")
                            name = name.strip()
                            location = location.strip()
                        return handler(name, "file", location)
                
                elif operation == "delete":
                    name = command.replace("delete", "").strip()
                    location = "downloads"
                    if " from " in name:
                        name, location = name.split(" from ")
                        name = name.strip()
                        location = location.strip()
                    return handler(name, location)
                
                elif operation == "empty":
                    location = command.replace("empty", "").strip()
                    return handler(location)
                
                elif operation == "shortcut":
                    if " for " in command:
                        target, shortcut = command.replace("create shortcut", "").split(" for ")
                        target = target.strip()
                        shortcut = shortcut.strip()
                        location = "desktop"
                        if " on " in shortcut:
                            shortcut, location = shortcut.split(" on ")
                            shortcut = shortcut.strip()
                            location = location.strip()
                        return handler(target, shortcut, location)
        
        speak("I'm not sure what file operation you want to perform. Try saying 'help' for available commands.")
        return False

# Initialize file manager
file_manager = FileManager()

# Update main command processing
def process_command(command):
    if not command:
        return
    
    print(f"Processing command: {command}")
    command = command.lower()
    
    # Handle file operations
    if any(op in command for op in ["create", "delete", "move", "copy", "rename", "share", "download", "empty", "shortcut"]):
        file_manager.handle_command(command)
        return
    
    # ... rest of existing command processing ...

# ---------- MAIN ----------
def get_available_commands():
    """Return a list of available commands for user reference"""
    return [
        "Media: 'play [song] on youtube', 'play [song] on spotify'",
        "Apps: 'open cursor', 'open vscode', 'open pictures', 'open downloads'",
        "System: 'volume [0-100]', 'brightness [0-100]', 'system info'",
        "Windows: 'focus window [name]', 'maximize window [name]', 'minimize window [name]'",
        "Mouse: 'move mouse [x] [y]', 'click', 'right click', 'scroll [amount]'",
        "Keyboard: 'type [text]', 'press key [key]', 'hotkey [key1] [key2]'",
        "Other: 'screenshot', 'network info', 'shutdown', 'restart'"
    ]

def handle_silence():
    """Handle when user doesn't respond after choosing voice mode"""
    speak("I'm still listening. Here are some things you can try:")
    commands = get_available_commands()
    for cmd in commands:
        speak(cmd)
        time.sleep(0.5)  # Pause between commands
    speak("Just say what you'd like me to do, or say 'text mode' to switch to typing.")

def main():
    try:
        print("\n" + "="*50)
        print("Welcome to KATE Assistant!")
        print("="*50 + "\n")

        # Step 1: Check and initialize voice
        print("Setting up voice...")
        global engine
        engine = setup_voice()
        if not engine:
            print("Voice engine not available. Will use text output only.")
        else:
            print("Voice setup complete")

        # Step 2: Check microphone
        print("\nChecking microphone...")
        try:
            mic = sr.Microphone()
            with mic as source:
                print("Adjusting for ambient noise...")
                recognizer = sr.Recognizer()
                recognizer.adjust_for_ambient_noise(source, duration=2)
            print("Microphone setup complete")
        except Exception as e:
            print(f"Microphone error: {str(e)}")
            print("Please check your microphone connection and try again.")
            return

        # Step 3: Authentication
        print("\nAuthentication required")
        try:
            authenticate()
        except Exception as e:
            print(f"Authentication failed: {str(e)}")
            return

        # Step 4: Start listening
        print("\nKATE is now active!")
        print("Say 'Hey Kate' to get started")
        print("Press Ctrl+C to exit\n")

        wake_detector = WakeWordDetector()
        wake_detector.start_listening()

        # Main loop
        while True:
            try:
                time.sleep(0.1)  # Reduce CPU usage
            except KeyboardInterrupt:
                print("\n\nGoodbye! KATE Assistant is shutting down...")
                wake_detector.stop_listening()
                break
            except Exception as e:
                print(f"\nError in main loop: {str(e)}")
                print("Attempting to continue...")
                continue

    except KeyboardInterrupt:
        print("\n\nGoodbye! KATE Assistant is shutting down...")
    except Exception as e:
        print(f"\nCritical error: {str(e)}")
        print("Please restart KATE Assistant.")

if __name__ == "__main__":
    main()
