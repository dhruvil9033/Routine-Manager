import os
import json
import speech_recognition as sr
import pyttsx3
import pyautogui
from dotenv import load_dotenv

load_dotenv()  # Load API keys from .env

class VoiceAssistant:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.r = sr.Recognizer()
        self.load_config()

    def load_config(self):
        with open('admin_whitelist.json') as f:
            self.security = json.load(f)

    def speak(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

    def listen(self):
        with sr.Microphone() as source:
            print("[Listening...]")
            audio = self.r.listen(source, timeout=5)
            try:
                return self.r.recognize_google(audio).lower()
            except:
                return ""

    def execute(self, command):
        # Security check for dangerous commands
        if any(cmd in command for cmd in self.security["dangerous_commands"]):
            self.speak("Security alert! Unauthorized command blocked.")
            return

        # File System
        if "list files" in command:
            files = os.listdir("C:\\Users\\Public")
            self.speak(f"Found {len(files)} files in Public folder")

        # GUI Control
        elif "open notepad" in command:
            os.system("notepad.exe")
            self.speak("Notepad opened")

        # Custom GPT-4 integration
        elif "ask ai" in command:
            query = command.replace("ask ai", "").strip()
            response = "GPT-4 response placeholder"  # Replace with actual API call
            self.speak(response)

        else:
            self.speak("Command not recognized")

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.speak("Voice assistant ready")
    while True:
        cmd = assistant.listen()
        if cmd:
            print(f"Command: {cmd}")
            assistant.execute(cmd)
        if "exit" in cmd:
            assistant.speak("Goodbye!")
            break