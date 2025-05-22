import speech_recognition as sr
import pyttsx3
import time
import os
import webbrowser
import datetime
import psutil
import pyautogui
import keyboard
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import shutil
import requests

class VoiceAssistant:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.is_authenticated = False
        self.is_sleeping = False
        self.sleep_word = "go to sleep"
        self.password = "40110825"
        self.driver = None
        self.current_note = None
        self.note_file = None
        self.is_listening = True
        self.background_listening = True
        self.setup_voice()
        
        # System paths and common applications
        self.system_paths = {
            'notepad': r'C:\Windows\System32\notepad.exe',
            'settings': r'ms-settings:',
            'downloads': os.path.expanduser('~\\Downloads'),
            'documents': os.path.expanduser('~\\Documents'),
            'pictures': os.path.expanduser('~\\Pictures'),
            'drivers': r'C:\Windows\System32\drivers',
            'system32': r'C:\Windows\System32',
            'program files': r'C:\Program Files',
            'program files (x86)': r'C:\Program Files (x86)'
        }
        
        # Common applications with fuzzy matching
        self.app_paths = {
            'whatsapp': [
                r'C:\Users\{}\AppData\Local\WhatsApp\WhatsApp.exe',
                r'C:\Program Files\WindowsApps\5319275A.WhatsAppDesktop_2.2405.4.0_x64__cv1g1gvanyjgm\WhatsApp.exe',
                r'C:\Program Files (x86)\WhatsApp\WhatsApp.exe'
            ],
            'vscode': r'C:\Users\{}\AppData\Local\Programs\VS Code\Code.exe',
            'mail': r'C:\Program Files\WindowsApps\microsoft.windowscommunicationsapps_16005.14326.21516.0_x64__8wekyb3d8bbwe\HxOutlook.exe',
            'chrome': r'C:\Program Files\Google\Chrome\Application\chrome.exe',
            'word': r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
            'excel': r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE',
            'powerpoint': r'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE',
            'outlook': r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
            'cursor': r'C:\Users\{}\AppData\Local\Programs\Cursor\Cursor.exe'
        }

    def setup_voice(self):
        voices = self.engine.getProperty('voices')
        for voice in voices:
            if "female" in voice.name.lower():
                self.engine.setProperty('voice', voice.id)
                break
        self.engine.setProperty('rate', 150)
        self.engine.setProperty('volume', 1.0)

    def setup_selenium(self):
        if self.driver is None:
            try:
                options = webdriver.ChromeOptions()
                options.add_argument('--start-maximized')
                self.driver = webdriver.Chrome(options=options)
            except:
                self.speak("Note: Chrome automation is not available")

    def speak(self, text):
        print(f"KATE: {text}")
        self.engine.say(text)
        self.engine.runAndWait()

    def listen(self):
        with self.microphone as source:
            print("Listening...")
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            audio = self.recognizer.listen(source)
            try:
                text = self.recognizer.recognize_google(audio)
                print(f"You said: {text}")
                return text.lower()
            except sr.UnknownValueError:
                return ""
            except sr.RequestError:
                self.speak("Sorry, I couldn't connect to the speech service")
                return ""

    def authenticate(self):
        self.speak("Welcome to KATE Assistant. Please enter the passcode.")
        print("\nYou can either:")
        print("1. Say the passcode (voice input)")
        print("2. Type the passcode (recommended for faster entry)")
        print("\nI highly recommend using text input (option 2) for faster and more reliable password entry.")
        print("Enter your choice (1 or 2): ", end='')
        
        choice = input().strip()
        attempts = 0
        voice_attempts = 0
        max_voice_attempts = 2
        
        while attempts < 3:
            if choice == "1":
                self.speak("Please say the passcode now.")
                text = self.listen()
                
                if not text:  # If voice recognition failed
                    voice_attempts += 1
                    if voice_attempts >= max_voice_attempts:
                        self.speak("I'm having trouble understanding your voice. Would you like to switch to text input? It's faster and more reliable.")
                        print("\nWould you like to switch to text input? (yes/no): ", end='')
                        switch_choice = input().strip().lower()
                        if switch_choice in ['yes', 'y']:
                            choice = "2"
                            continue
                        else:
                            self.speak("Please try saying the passcode again.")
                            continue
                    else:
                        self.speak("I didn't catch that. Please try again.")
                        continue
            else:
                print("Please type the passcode: ", end='')
                text = input().strip().lower()
            
            if text == self.password:
                self.is_authenticated = True
                self.speak("Authentication successful. How can I help you?")
                return True
            else:
                attempts += 1
                if attempts < 3:
                    self.speak(f"Invalid passcode. {3-attempts} attempts remaining.")
                    print(f"\nInvalid passcode. {3-attempts} attempts remaining.")
                    if choice == "1":
                        print("\nWould you like to switch to text input? It's faster and more reliable. (yes/no): ", end='')
                        switch_choice = input().strip().lower()
                        if switch_choice in ['yes', 'y']:
                            choice = "2"
                else:
                    self.speak("This assistant is only for Wallace. Goodbye.")
                    print("\nThis assistant is only for Wallace. Goodbye.")
                    return False

    def fuzzy_match_app(self, app_name):
        """Fuzzy match application names for natural language input"""
        app_name = app_name.lower()
        best_match = None
        best_score = 0
        
        for known_app in self.app_paths.keys():
            # Simple matching for now, can be enhanced with fuzzywuzzy if needed
            if app_name in known_app or known_app in app_name:
                return known_app
        return None

    def open_application(self, app_name):
        try:
            app_name = app_name.lower()
            
            # Handle special system locations
            if app_name in self.system_paths:
                path = self.system_paths[app_name]
                if os.path.isdir(path):
                    os.startfile(path)
                    self.speak(f"Opening {app_name} folder")
                else:
                    os.startfile(path)
                    self.speak(f"Opening {app_name}")
                return
            
            # Try fuzzy matching for applications
            matched_app = self.fuzzy_match_app(app_name)
            if matched_app:
                # Special handling for WhatsApp
                if matched_app == 'whatsapp':
                    for path_template in self.app_paths['whatsapp']:
                        try:
                            path = path_template.format(os.getenv('USERNAME'))
                            if os.path.exists(path):
                                os.startfile(path)
                                self.speak("Opening WhatsApp")
                                return
                        except:
                            continue
                    # If no path worked, try opening WhatsApp Web
                    webbrowser.open('https://web.whatsapp.com')
                    self.speak("Opening WhatsApp Web")
                    return
                
                # Handle other applications
                path = self.app_paths[matched_app].format(os.getenv('USERNAME'))
                try:
                    os.startfile(path)
                    self.speak(f"Opening {matched_app}")
                    return
                except:
                    pass
            
            # Try direct application name
            try:
                os.startfile(app_name)
                self.speak(f"Opening {app_name}")
            except:
                self.speak(f"I couldn't find {app_name}. Please try being more specific.")

        except Exception as e:
            self.speak(f"Could not open {app_name}. Error: {str(e)}")

    def play_media(self, query, platform="youtube"):
        if platform == "youtube":
            webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
            self.speak(f"Searching YouTube for {query}")
        elif platform == "spotify":
            webbrowser.open(f"https://open.spotify.com/search/{query}")
            self.speak(f"Searching Spotify for {query}")

    def search_web(self, query, platform="google"):
        try:
            if platform == "google":
                webbrowser.open(f"https://www.google.com/search?q={query}")
                self.speak(f"Searching Google for {query}")
            elif platform == "chatgpt":
                if self.driver is None:
                    self.setup_selenium()
                if self.driver:
                    self.driver.get("https://chat.openai.com")
                    chat_input = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "textarea"))
                    )
                    chat_input.clear()
                    chat_input.send_keys(query)
                    chat_input.send_keys(Keys.RETURN)
                    self.speak("Searching ChatGPT for your query")
                else:
                    webbrowser.open("https://chat.openai.com")
                    self.speak("Please type your query in ChatGPT")
        except Exception as e:
            self.speak(f"Could not perform web search. Error: {str(e)}")

    def type_text(self, text, app="notepad"):
        try:
            # Open the application if not already open
            if app == "notepad":
                os.startfile(self.system_paths['notepad'])
            elif app == "word":
                os.startfile("winword.exe")
            
            # Wait for application to open
            time.sleep(1)
            
            # Copy text to clipboard
            pyperclip.copy(text)
            
            # Paste the text
            keyboard.press_and_release('ctrl+v')
            
            self.speak("Text has been typed")
        except Exception as e:
            self.speak(f"Could not type text. Error: {str(e)}")

    def get_system_info(self):
        try:
            cpu_percent = psutil.cpu_percent()
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            info = f"""
            CPU Usage: {cpu_percent}%
            Memory Usage: {memory.percent}%
            Disk Usage: {disk.percent}%
            """
            self.speak(info)
            
            # Additional system information
            self.speak("Would you like to know about specific system locations?")
            print("\nAvailable system locations:")
            for location in self.system_paths.keys():
                print(f"- {location}")
            print("\nYou can ask about any of these locations.")
            
        except Exception as e:
            self.speak(f"Could not get system information. Error: {str(e)}")

    def system_control(self, action):
        try:
            if "volume" in action:
                if "up" in action or "increase" in action:
                    for _ in range(3):
                        keyboard.press_and_release('volume up')
                    self.speak("Volume increased")
                elif "down" in action or "decrease" in action:
                    for _ in range(3):
                        keyboard.press_and_release('volume down')
                    self.speak("Volume decreased")
                elif "mute" in action:
                    keyboard.press_and_release('volume mute')
                    self.speak("Volume muted")
            
            elif "bluetooth" in action:
                if "on" in action or "enable" in action:
                    os.system('powershell -command "Start-Process ms-settings:bluetooth"')
                    self.speak("Opening Bluetooth settings")
                elif "off" in action or "disable" in action:
                    os.system('powershell -command "Start-Process ms-settings:bluetooth"')
                    self.speak("Opening Bluetooth settings to disable")
            
            elif "wifi" in action:
                if "on" in action or "enable" in action:
                    os.system('powershell -command "Start-Process ms-settings:network-wifi"')
                    self.speak("Opening WiFi settings")
                elif "off" in action or "disable" in action:
                    os.system('powershell -command "Start-Process ms-settings:network-wifi"')
                    self.speak("Opening WiFi settings to disable")
            
            elif "sleep" in action:
                os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')
                self.speak("Putting computer to sleep")
            
            elif "restart" in action:
                self.speak("Restarting computer in 10 seconds")
                os.system('shutdown /r /t 10')
            
            elif "shutdown" in action:
                self.speak("Shutting down computer in 10 seconds")
                os.system('shutdown /s /t 10')
            
            else:
                self.speak("I don't understand that system control command")

        except Exception as e:
            self.speak(f"Could not perform system control. Error: {str(e)}")

    def keyboard_shortcut(self, action):
        try:
            if "switch tab" in action or "next tab" in action:
                keyboard.press_and_release('ctrl+tab')
                self.speak("Switched to next tab")
            elif "previous tab" in action or "last tab" in action:
                keyboard.press_and_release('ctrl+shift+tab')
                self.speak("Switched to previous tab")
            elif "close tab" in action:
                keyboard.press_and_release('ctrl+w')
                self.speak("Closed current tab")
            elif "new tab" in action:
                keyboard.press_and_release('ctrl+t')
                self.speak("Opened new tab")
            elif "screenshot" in action:
                keyboard.press_and_release('win+shift+s')
                self.speak("Taking screenshot")
            elif "copy" in action:
                keyboard.press_and_release('ctrl+c')
                self.speak("Copied to clipboard")
            elif "paste" in action:
                keyboard.press_and_release('ctrl+v')
                self.speak("Pasted from clipboard")
            else:
                self.speak("I don't understand that keyboard shortcut")

        except Exception as e:
            self.speak(f"Could not perform keyboard shortcut. Error: {str(e)}")

    def start_note(self, editor="notepad"):
        try:
            # Open the text editor
            if editor == "notepad":
                os.startfile(self.system_paths['notepad'])
            elif editor == "word":
                os.startfile(self.app_paths['word'].format(os.getenv('USERNAME')))
            
            # Wait for editor to open
            time.sleep(2)
            
            # Initialize note tracking
            self.current_note = {
                'editor': editor,
                'start_time': datetime.datetime.now(),
                'is_listening': True
            }
            
            self.speak(f"Started new note in {editor}. I'm listening...")
            
        except Exception as e:
            self.speak(f"Could not start note. Error: {str(e)}")
            self.current_note = None

    def add_to_note(self, text):
        try:
            if self.current_note and self.current_note['is_listening']:
                # Copy text to clipboard
                pyperclip.copy(text)
                
                # Wait a moment
                time.sleep(0.5)
                
                # Paste into active window
                keyboard.press_and_release('ctrl+v')
                
                # Add new line
                keyboard.press_and_release('enter')
                
                self.speak("Added to note")
                
        except Exception as e:
            self.speak(f"Could not add to note. Error: {str(e)}")

    def end_note(self):
        try:
            if self.current_note:
                # Save the document
                keyboard.press_and_release('ctrl+s')
                
                # Clear note tracking
                self.current_note = None
                self.speak("Note saved and closed")
                
        except Exception as e:
            self.speak(f"Could not end note. Error: {str(e)}")

    def suggest_commands(self, command):
        """Suggest similar commands when a command is not understood"""
        command = command.lower()
        suggestions = []
        
        # Check for similar application names
        for app in self.app_paths.keys():
            if app in command or command in app:
                suggestions.append(f"open {app}")
        
        # Check for similar system locations
        for location in self.system_paths.keys():
            if location in command or command in location:
                suggestions.append(f"open {location}")
        
        # Check for common command patterns
        if "note" in command:
            suggestions.extend(["start taking notes", "start taking notes in word", "stop taking notes"])
        if "play" in command:
            suggestions.extend(["play [song] on youtube", "play [song] on spotify"])
        if "search" in command:
            suggestions.extend(["search [query]", "search [query] in chatgpt"])
        if "volume" in command:
            suggestions.extend(["volume up", "volume down", "mute"])
        if "bluetooth" in command or "wifi" in command:
            suggestions.extend(["turn bluetooth on", "turn bluetooth off", "turn wifi on", "turn wifi off"])
        
        return suggestions

    def web_control(self, action):
        try:
            if "refresh" in action or "reload" in action:
                keyboard.press_and_release('f5')
                self.speak("Refreshing page")
            
            elif "clear" in action or "delete" in action:
                # Clear the current input/address bar
                keyboard.press_and_release('ctrl+l')  # Focus on address bar
                time.sleep(0.5)
                keyboard.press_and_release('ctrl+a')  # Select all
                time.sleep(0.5)
                keyboard.press_and_release('backspace')  # Delete
                self.speak("Cleared current text")
            
            elif "write" in action or "type" in action:
                # Extract text to write
                text = action.replace("write", "").replace("type", "").strip()
                if text:
                    pyperclip.copy(text)
                    time.sleep(0.5)
                    keyboard.press_and_release('ctrl+v')
                    self.speak(f"Typed: {text}")
                else:
                    self.speak("What would you like me to type?")
            
            elif "search" in action:
                # Extract search query
                query = action.replace("search", "").strip()
                if query:
                    pyperclip.copy(query)
                    time.sleep(0.5)
                    keyboard.press_and_release('ctrl+v')
                    keyboard.press('enter')
                    self.speak(f"Searching for: {query}")
                else:
                    self.speak("What would you like me to search for?")
            
            else:
                self.speak("I can refresh the page, clear text, or type new text. What would you like me to do?")

        except Exception as e:
            self.speak(f"Could not perform web control action. Error: {str(e)}")

    def process_command(self, command):
        if not command:
            return
            
        print(f"Processing command: {command}")
        command = command.lower()
        
        # Handle web control commands
        if any(word in command for word in ["refresh", "reload", "clear", "delete", "write", "type", "search"]):
            self.web_control(command)
            return

        # Handle search commands
        if "search" in command:
            # Extract search query and platform
            query = command.replace("search", "").strip()
            platform = "google"  # Default to Google
            
            # Check for specific platform
            if " on " in query:
                query, platform = query.split(" on ", 1)
                platform = platform.strip()
            
            # Handle different platforms
            if platform == "google":
                search_url = f"https://www.google.com/search?q={'+'.join(query.split())}"
                webbrowser.open(search_url)
                self.speak(f"Searching for {query} on Google")
            elif platform == "chatgpt":
                webbrowser.open("https://chat.openai.com")
                time.sleep(2)  # Wait for page to load
                pyperclip.copy(query)
                keyboard.press_and_release('ctrl+v')
                keyboard.press('enter')
                self.speak(f"Searching for {query} on ChatGPT")
            else:
                self.speak(f"I can search on Google or ChatGPT. Would you like me to search for {query} on Google?")
            return

        if self.sleep_word in command:
            self.is_sleeping = True
            self.speak("Going to sleep. Say 'wake up' when you need me.")
            return

        if self.is_sleeping and "wake up" in command:
            self.is_sleeping = False
            self.speak("I'm awake. How can I help you?")
            return

        if not self.is_sleeping:
            # Handle background listening commands
            if command.lower() == "okay":
                self.is_listening = True
                self.speak("I'm listening")
                return
            elif command.lower() == "stop":
                self.is_listening = False
                self.speak("Paused listening")
                return
            elif command.lower() == "continue":
                self.is_listening = True
                self.speak("Resumed listening")
                return

            if not self.is_listening:
                return

            # Check for the special shutdown command
            if "i'm tired" in command.lower() or "im tired" in command.lower():
                current_hour = datetime.datetime.now().hour
                if current_hour < 24:  # Before midnight
                    self.speak("You're tired bro? It's not even midnight yet! You're weird...")
                else:  # After midnight
                    self.speak("Damn man, you're one heck of a pain in the ass dude! See you tomorrow to make the world a better place!")
                self.speak("Shutting down in 10 seconds...")
                os.system('shutdown /s /t 10')
                return

            # Note taking commands
            if "start taking notes" in command or "start note" in command:
                editor = "notepad"  # Default
                if "in word" in command:
                    editor = "word"
                self.start_note(editor)
            elif "end note" in command or "stop taking notes" in command:
                self.end_note()
            elif self.current_note and self.current_note['is_listening']:
                self.add_to_note(command)
                return

            # System control commands
            if any(word in command for word in ["volume", "bluetooth", "wifi", "sleep", "restart", "shutdown"]):
                self.system_control(command)
                return

            # Keyboard shortcut commands
            if any(phrase in command for phrase in ["switch tab", "close tab", "new tab", "screenshot", "copy", "paste"]):
                self.keyboard_shortcut(command)
                return

            # File and application commands
            if "open" in command:
                self.open_application(command.replace("open", "").strip())
            
            elif "play" in command:
                if "on youtube" in command:
                    query = command.replace("play", "").replace("on youtube", "").strip()
                    self.play_media(query, "youtube")
                elif "on spotify" in command:
                    query = command.replace("play", "").replace("on spotify", "").strip()
                    self.play_media(query, "spotify")
            
            elif "help" in command or "what can you do" in command:
                self.speak("""
                I can help you with:
                - Opening applications (WhatsApp, VS Code, Mail, Office apps)
                - Opening system locations (Downloads, Documents, Pictures)
                - System controls (volume, bluetooth, wifi, sleep, restart)
                - Keyboard shortcuts (switch tabs, screenshots, copy/paste)
                - Taking notes in Notepad or Word (start taking notes, stop taking notes)
                - Playing media (YouTube, Spotify)
                - Searching the web (Google, ChatGPT)
                - Special commands (try saying "I'm tired" to see what happens!)
                Just tell me what you'd like me to do.
                """)
            
            else:
                suggestions = self.suggest_commands(command)
                if suggestions:
                    self.speak("I'm not sure what you want me to do. Did you mean:")
                    for suggestion in suggestions[:3]:  # Show top 3 suggestions
                        self.speak(f"- {suggestion}")
                else:
                    self.speak("I'm not sure what you want me to do. Try saying 'help' to see what I can do.")

    def run(self):
        try:
            if not self.authenticate():
                return
            
            self.speak("I'm ready to help. Just tell me what you need.")
            
            while True:
                if self.background_listening:
                    command = self.listen()
                    if command:
                        self.process_command(command)
                
        except Exception as e:
            print(f"Error: {str(e)}")
            self.speak("I encountered an error. Please restart me.")

# ---------- FILE MANAGEMENT ----------
class FileManager:
    def __init__(self):
        self.user_home = os.path.expanduser("~")
        self.common_paths = {
            "desktop": os.path.join(self.user_home, "Desktop"),
            "downloads": os.path.join(self.user_home, "Downloads"),
            "documents": os.path.join(self.user_home, "Documents"),
            "pictures": os.path.join(self.user_home, "Pictures")
        }
        self.whatsapp_contacts = {}
        self.load_contacts()

    def load_contacts(self):
        try:
            with open("whatsapp_contacts.json", "r") as f:
                self.whatsapp_contacts = json.load(f)
        except:
            self.whatsapp_contacts = {}

    def save_contacts(self):
        with open("whatsapp_contacts.json", "w") as f:
            json.dump(self.whatsapp_contacts, f)

    def create_folder(self, folder_name, location="downloads"):
        try:
            if location in self.common_paths:
                path = os.path.join(self.common_paths[location], folder_name)
                os.makedirs(path, exist_ok=True)
                self.speak(f"Created folder {folder_name} in {location}")
                return True
            return False
        except Exception as e:
            self.speak(f"Could not create folder: {str(e)}")
            return False

    def create_file(self, file_name, content="", location="downloads"):
        try:
            if location in self.common_paths:
                path = os.path.join(self.common_paths[location], file_name)
                with open(path, "w") as f:
                    f.write(content)
                self.speak(f"Created file {file_name} in {location}")
                return True
            return False
        except Exception as e:
            self.speak(f"Could not create file: {str(e)}")
            return False

    def delete_item(self, item_name, location="downloads"):
        try:
            if location in self.common_paths:
                path = os.path.join(self.common_paths[location], item_name)
                if os.path.exists(path):
                    if os.path.isfile(path):
                        os.remove(path)
                    else:
                        shutil.rmtree(path)
                    self.speak(f"Deleted {item_name} from {location}")
                    return True
            return False
        except Exception as e:
            self.speak(f"Could not delete item: {str(e)}")
            return False

    def share_file(self, file_name, contact_name, platform="whatsapp"):
        try:
            if platform == "whatsapp":
                if contact_name in self.whatsapp_contacts:
                    # Open WhatsApp Web
                    webbrowser.open("https://web.whatsapp.com")
                    time.sleep(5)  # Wait for WhatsApp Web to load
                    
                    # Search for contact
                    pyperclip.copy(contact_name)
                    keyboard.press_and_release('ctrl+f')
                    time.sleep(1)
                    keyboard.press_and_release('ctrl+v')
                    time.sleep(2)
                    keyboard.press('enter')
                    time.sleep(1)
                    
                    # Click attachment button
                    pyautogui.click(x=1000, y=700)  # Adjust coordinates as needed
                    time.sleep(1)
                    
                    # Click document
                    pyautogui.click(x=1000, y=600)  # Adjust coordinates as needed
                    time.sleep(1)
                    
                    # Select file
                    file_path = os.path.join(self.common_paths["downloads"], file_name)
                    pyperclip.copy(file_path)
                    keyboard.press_and_release('ctrl+v')
                    time.sleep(1)
                    keyboard.press('enter')
                    
                    self.speak(f"Sharing {file_name} with {contact_name} on WhatsApp")
                    return True
                else:
                    self.speak(f"Contact {contact_name} not found. Would you like to add them?")
                    return False
            elif platform == "email":
                # Email sharing logic here
                pass
            return False
        except Exception as e:
            self.speak(f"Could not share file: {str(e)}")
            return False

    def download_file(self, url, file_name, location="downloads"):
        try:
            if location in self.common_paths:
                path = os.path.join(self.common_paths[location], file_name)
                response = requests.get(url, stream=True)
                with open(path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                self.speak(f"Downloaded {file_name} to {location}")
                return True
            return False
        except Exception as e:
            self.speak(f"Could not download file: {str(e)}")
            return False

# ---------- ENHANCED BROWSER MANAGEMENT ----------
class EnhancedBrowserManager:
    def __init__(self):
        self.file_manager = FileManager()

    def handle_download(self, url, file_name):
        self.speak(f"Where would you like to save {file_name}?")
        self.speak("Options: downloads, desktop, documents, pictures")
        location = input("Enter location: ").lower()
        if location in ["downloads", "desktop", "documents", "pictures"]:
            return self.file_manager.download_file(url, file_name, location)
        else:
            self.speak("Invalid location. Saving to downloads.")
            return self.file_manager.download_file(url, file_name, "downloads")

    def close_tab(self):
        keyboard.press_and_release('ctrl+w')
        self.speak("Closed current tab")

    def close_window(self):
        keyboard.press_and_release('alt+f4')
        self.speak("Closed current window")

# Initialize managers
file_manager = FileManager()
browser_manager = EnhancedBrowserManager()

# Update handle_command function to include new capabilities
def handle_command(cmd):
    # ... existing code ...
    
    # Handle file operations
    if "create folder" in cmd:
        folder_name = cmd.replace("create folder", "").strip()
        location = "downloads"  # Default location
        if " in " in folder_name:
            folder_name, location = folder_name.split(" in ")
            folder_name = folder_name.strip()
            location = location.strip()
        file_manager.create_folder(folder_name, location)
        return

    if "create file" in cmd:
        file_name = cmd.replace("create file", "").strip()
        location = "downloads"  # Default location
        if " in " in file_name:
            file_name, location = file_name.split(" in ")
            file_name = file_name.strip()
            location = location.strip()
        content = input("Enter file content (press Enter twice to finish):\n")
        file_manager.create_file(file_name, content, location)
        return

    if "delete" in cmd:
        item_name = cmd.replace("delete", "").strip()
        location = "downloads"  # Default location
        if " from " in item_name:
            item_name, location = item_name.split(" from ")
            item_name = item_name.strip()
            location = location.strip()
        file_manager.delete_item(item_name, location)
        return

    if "share" in cmd:
        if " with " in cmd:
            file_name, contact = cmd.replace("share", "").split(" with ")
            file_name = file_name.strip()
            contact = contact.strip()
            platform = "whatsapp"  # Default platform
            if " on " in contact:
                contact, platform = contact.split(" on ")
                contact = contact.strip()
                platform = platform.strip()
            file_manager.share_file(file_name, contact, platform)
        return

    # Handle browser operations
    if "close tab" in cmd:
        browser_manager.close_tab()
        return

    if "close window" in cmd:
        browser_manager.close_window()
        return

    # ... rest of existing code ...

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run() 