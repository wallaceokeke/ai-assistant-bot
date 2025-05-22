import json
import time
import random
from datetime import datetime
import sqlite3
from pathlib import Path
import psutil
import pyttsx3
import speech_recognition as sr
import re

class EmotionalIntelligence:
    def __init__(self):
        self.emotion_patterns = {
            'boredom': [
                'bored', 'boring', 'nothing to do', 'tired', 'dull',
                'uninteresting', 'monotonous', 'tedious'
            ],
            'happiness': [
                'happy', 'excited', 'great', 'wonderful', 'amazing',
                'fantastic', 'awesome', 'joy'
            ],
            'frustration': [
                'frustrated', 'annoyed', 'angry', 'mad', 'irritated',
                'upset', 'disappointed'
            ],
            'tiredness': [
                'tired', 'exhausted', 'sleepy', 'drowsy', 'fatigued',
                'worn out'
            ]
        }
        
        self.response_templates = {
            'boredom': [
                "Would you like to hear something interesting?",
                "I can tell you a fun fact!",
                "How about we try something new?",
                "Would you like to hear a joke?",
                "I know some interesting trivia if you're interested!"
            ],
            'happiness': [
                "That's wonderful to hear!",
                "I'm glad you're feeling good!",
                "Your positive energy is contagious!",
                "That's fantastic!",
                "I'm happy that you're happy!"
            ],
            'frustration': [
                "I understand this can be frustrating. Would you like to talk about it?",
                "I'm here to help if you need anything.",
                "Let's try to find a solution together.",
                "Would you like to take a short break?",
                "I can help you work through this."
            ],
            'tiredness': [
                "Would you like to take a short break?",
                "How about some energizing music?",
                "Maybe we should take it easy for a bit.",
                "Would you like to hear something uplifting?",
                "How about a quick stretch or walk?"
            ]
        }
        
        self.fun_facts = [
            "A day on Venus is longer than its year!",
            "Honey never spoils. Archaeologists have found pots of honey in ancient Egyptian tombs that are over 3,000 years old!",
            "The shortest war in history was between Britain and Zanzibar on August 27, 1896. Zanzibar surrendered after just 38 minutes.",
            "A group of flamingos is called a 'flamboyance'.",
            "The first oranges weren't orange - they were green!",
            "The average person spends 6 months of their lifetime waiting for red lights to turn green.",
            "A jiffy is an actual unit of time: 1/100th of a second.",
            "The first computer programmer was a woman named Ada Lovelace.",
            "The first computer bug was an actual bug - a moth found in a computer in 1947.",
            "The first computer mouse was made of wood!"
        ]
        
        self.jokes = [
            "Why don't scientists trust atoms? Because they make up everything!",
            "Why did the scarecrow win an award? Because he was outstanding in his field!",
            "What do you call a fake noodle? An impasta!",
            "How does a penguin build its house? Igloos it together!",
            "Why don't eggs tell jokes? They'd crack each other up!",
            "What do you call a can opener that doesn't work? A can't opener!",
            "Why did the math book look so sad? Because it had too many problems!",
            "What do you call a bear with no teeth? A gummy bear!",
            "Why don't skeletons fight each other? They don't have the guts!",
            "What do you call a fish with no eyes? Fsh!"
        ]

    def detect_emotion(self, text):
        text = text.lower()
        for emotion, patterns in self.emotion_patterns.items():
            if any(pattern in text for pattern in patterns):
                return emotion
        return None

    def get_emotional_response(self, emotion):
        if emotion in self.response_templates:
            return random.choice(self.response_templates[emotion])
        return None

    def get_fun_fact(self):
        return random.choice(self.fun_facts)

    def get_joke(self):
        return random.choice(self.jokes)

class ActivityTracker:
    def __init__(self):
        self.db_path = "user_activity.db"
        self.init_db()
        self.current_session = {
            'start_time': datetime.now(),
            'activities': [],
            'apps_opened': set(),
            'patterns': {}
        }

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS activities
                    (timestamp TEXT, activity TEXT, app_name TEXT, duration INTEGER)''')
        c.execute('''CREATE TABLE IF NOT EXISTS patterns
                    (pattern TEXT, intent TEXT, time_of_day TEXT, frequency INTEGER)''')
        conn.commit()
        conn.close()

    def log_activity(self, activity, app_name=None):
        timestamp = datetime.now()
        self.current_session['activities'].append({
            'timestamp': timestamp,
            'activity': activity,
            'app_name': app_name
        })
        
        if app_name:
            self.current_session['apps_opened'].add(app_name)
        
        # Save to database
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("INSERT INTO activities VALUES (?, ?, ?, ?)",
                 (timestamp.isoformat(), activity, app_name, 0))
        conn.commit()
        conn.close()

    def get_recent_activities(self, limit=5):
        return self.current_session['activities'][-limit:]

    def summarize_activities(self):
        if not self.current_session['activities']:
            return "You haven't done anything yet in this session."
        
        summary = "Here's what you've been up to:\n"
        for activity in self.current_session['activities']:
            summary += f"- {activity['activity']}"
            if activity['app_name']:
                summary += f" in {activity['app_name']}"
            summary += "\n"
        return summary

    def learn_pattern(self, activity, intent, time_of_day):
        pattern_key = f"{activity}_{intent}_{time_of_day}"
        if pattern_key in self.current_session['patterns']:
            self.current_session['patterns'][pattern_key] += 1
        else:
            self.current_session['patterns'][pattern_key] = 1

class EnhancedKATE:
    def __init__(self):
        self.emotional_intelligence = EmotionalIntelligence()
        self.activity_tracker = ActivityTracker()
        self.setup_voice()
        self.setup_speech_recognition()
        
        # Time-based suggestions
        self.time_suggestions = {
            'morning': [
                "Would you like to start your day with some energizing music?",
                "How about checking your schedule for today?",
                "Would you like to hear the weather forecast?"
            ],
            'afternoon': [
                "Time for a productivity playlist?",
                "Would you like to take a short break?",
                "How about some background music while you work?"
            ],
            'evening': [
                "Would you like some relaxing music?",
                "How about a summary of your day?",
                "Would you like to plan for tomorrow?"
            ]
        }

    def setup_voice(self):
        try:
            self.engine = pyttsx3.init()
            voices = self.engine.getProperty('voices')
            self.engine.setProperty('voice', voices[0].id)
            self.engine.setProperty('rate', 150)
        except Exception as e:
            print(f"Voice setup error: {e}")
            self.engine = None

    def setup_speech_recognition(self):
        try:
            self.recognizer = sr.Recognizer()
            self.microphone = sr.Microphone()
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source)
        except Exception as e:
            print(f"Speech recognition setup error: {e}")
            self.recognizer = None
            self.microphone = None

    def speak(self, text):
        print(f"KATE: {text}")
        if self.engine:
            self.engine.say(text)
            self.engine.runAndWait()

    def get_time_of_day(self):
        hour = datetime.now().hour
        if 5 <= hour < 12:
            return 'morning'
        elif 12 <= hour < 17:
            return 'afternoon'
        else:
            return 'evening'

    def get_time_based_suggestion(self):
        time_of_day = self.get_time_of_day()
        return random.choice(self.time_suggestions[time_of_day])

    def handle_emotion(self, text):
        emotion = self.emotional_intelligence.detect_emotion(text)
        if emotion:
            response = self.emotional_intelligence.get_emotional_response(emotion)
            if response:
                self.speak(response)
                if emotion == 'boredom':
                    if random.random() < 0.5:
                        self.speak(self.emotional_intelligence.get_fun_fact())
                    else:
                        self.speak(self.emotional_intelligence.get_joke())

    def handle_activity_summary(self, text):
        if any(word in text.lower() for word in ['what did i do', 'what have i done', 'my activities', 'recent activity']):
            self.speak(self.activity_tracker.summarize_activities())

    def process_command(self, text):
        # Log the interaction
        self.activity_tracker.log_activity(f"User said: {text}")
        
        # Handle emotions
        self.handle_emotion(text)
        
        # Handle activity summary requests
        self.handle_activity_summary(text)
        
        # Handle basic commands
        text = text.lower()
        if 'open' in text:
            app_name = text.replace('open', '').strip()
            self.speak(f"I'll open {app_name} for you.")
            self.activity_tracker.log_activity(f"Opened {app_name}", app_name)
            
            # Learn pattern
            time_of_day = self.get_time_of_day()
            self.activity_tracker.learn_pattern(app_name, 'open', time_of_day)
            
            # Make time-based suggestion
            if app_name.lower() in ['vscode', 'code', 'visual studio']:
                self.speak("Would you like me to play some coding music while you work?")
        
        elif 'no' in text or 'nope' in text or 'nah' in text:
            self.speak("That's okay! Let me know if you need anything else.")
        
        elif 'yes' in text or 'sure' in text or 'okay' in text:
            self.speak("Great! I'll do that for you.")
        
        else:
            # If no specific command is recognized, offer a time-based suggestion
            self.speak(self.get_time_based_suggestion())

    def run(self):
        self.speak("Hello! I'm KATE, your enhanced assistant. How can I help you today?")
        
        while True:
            try:
                if self.microphone and self.recognizer:
                    with self.microphone as source:
                        print("Listening...")
                        audio = self.recognizer.listen(source)
                        try:
                            text = self.recognizer.recognize_google(audio)
                            print(f"You said: {text}")
                            self.process_command(text)
                        except sr.UnknownValueError:
                            print("Could not understand audio")
                        except sr.RequestError as e:
                            print(f"Could not request results; {e}")
                else:
                    text = input("Type your command: ")
                    self.process_command(text)
                    
            except KeyboardInterrupt:
                self.speak("Goodbye! Have a great day!")
                break
            except Exception as e:
                print(f"Error: {e}")
                continue

if __name__ == "__main__":
    kate = EnhancedKATE()
    kate.run() 