import os
import json
import subprocess
import speech_recognition as sr
import pyttsx3
from datetime import datetime
import ctypes
import sys
import difflib
import win32com.client
from routine_gui import RoutineManagerGUI


class CompleteSystemController:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.r = sr.Recognizer()
        self.memory_file = "app_paths.json"
        self.learned_apps = self.load_memory()
        self.routine_file = "routines.json"
        self.routines = self.load_routines()

        self.system_commands = {
            "file explorer": ("explorer.exe", False),
            "control panel": ("control.exe", False),
            "task manager": ("taskmgr.exe", True),
            "cmd": ("cmd.exe", True),
            "powershell": ("powershell.exe", True),
            "notepad": ("notepad.exe", False),
            "registry": ("regedit.exe", True)
        }
    
    def load_routines(self):
        try:
            with open(self.routine_file, 'r') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def save_routines(self):
        with open(self.routine_file, 'w') as f:
            json.dump(self.routines, f, indent=2)

    def create_routine(self):
        self.speak("What's the name of the routine?")
        name = input("Routine name: ").strip().lower()
        apps = []

        while True:
            self.speak("Enter app name to add (or type 'done'):")
            app_name = input("App name (or 'done'): ").strip().lower()
            if app_name == "done":
                break
            admin_input = input(f"Should {app_name} be run as admin? (yes/no): ").strip().lower()
            apps.append({
                "name": app_name,
                "admin": admin_input in ["yes", "y"]
            })

        self.routines[name] = apps
        self.save_routines()
        self.speak(f"Routine '{name}' saved.")

    def list_routines(self):
        if not self.routines:
            self.speak("No routines found.")
        else:
            self.speak("Here are your routines:")
            for name in self.routines:
                print(f"- {name}")

    def run_routine(self, name):
        routine = self.routines.get(name)
        if not routine:
            self.speak(f"No routine named {name} found.")
            return

        self.speak(f"Starting routine: {name}")
        for app in routine:
            self.open_app(app["name"], admin=app["admin"])


    def load_memory(self):
        try:
            with open(self.memory_file, 'r') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def save_memory(self):
        with open(self.memory_file, 'w') as f:
            json.dump(self.learned_apps, f, indent=2)

    def speak(self, text):
        print(f"ASSISTANT: {text}")
        self.engine.say(text)
        self.engine.runAndWait()

    def listen(self):
        with sr.Microphone() as source:
            print("\n[Listening...]")
            try:
                audio = self.r.listen(source, timeout=3)
                return self.r.recognize_google(audio).lower()
            except sr.WaitTimeoutError:
                self.speak("Listening timed out, please try again.")
                return ""
            except Exception as e:
                print(f"Recognition error: {e}")
                return ""

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def run_as_admin(self, command):
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", command, None, None, 1)
            return True
        except Exception as e:
            print(f"Admin elevation failed: {e}")
            return False

    def fuzzy_match(self, app_name, known_names):
        match = difflib.get_close_matches(app_name, known_names, n=1, cutoff=0.7)
        return match[0] if match else app_name

    def resolve_shortcut_name(self, shortcut_path):
        try:
            return os.path.splitext(os.path.basename(shortcut_path))[0]
        except:
            return None

    def find_app_path(self, app_name):
        app_name_lower = app_name.lower()

        if app_name_lower in self.system_commands:
            return self.system_commands[app_name_lower]

        if app_name_lower in self.learned_apps:
            path = self.learned_apps[app_name_lower]["path"]
            if os.path.exists(path):
                return (path, self.learned_apps[app_name_lower].get("requires_admin", False))

        search_paths = [
            r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs",
            os.path.expandvars(r"%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs"),
            os.path.expandvars(r"%USERPROFILE%\\Desktop"),
            os.path.expandvars(r"%PROGRAMFILES%"),
            os.path.expandvars(r"%PROGRAMFILES(X86)%"),
            os.path.expandvars(r"%LOCALAPPDATA%")
        ]

        matches = []
        for base_path in search_paths:
            for root, _, files in os.walk(base_path):
                for file in files:
                    if file.lower().endswith(('.lnk', '.exe')):
                        display_name = self.resolve_shortcut_name(os.path.join(root, file)).lower()
                        if app_name_lower in display_name:
                            matches.append(os.path.join(root, file))

        if not matches:
            return (None, False)

        matches = sorted(matches, key=lambda m: 0 if m.endswith('.lnk') else 1)

        if len(matches) == 1:
            return (matches[0], False)

        self.speak("I found multiple matches. Please say the option number.")
        for idx, path in enumerate(matches[:5], 1):
            name = self.resolve_shortcut_name(path)
            self.speak(f"Option {idx}: {name}")

        response = self.listen()
        option_number = None

        for word in response.split():
            if word.isdigit():
                option_number = int(word)
                break
            word_to_number = {
                "one": 1, "two": 2, "three": 3, "four": 4, "five": 5
            }
            if word in word_to_number:
                option_number = word_to_number[word]
                break

        if option_number and 1 <= option_number <= len(matches):
            return (matches[option_number - 1], False)

        self.speak("I couldn't understand the option number. Please try again.")
        return (None, False)

    def open_app(self, app_name, admin=False):
        all_known = list(self.learned_apps.keys()) + list(self.system_commands.keys())
        app_name = self.fuzzy_match(app_name, all_known)
        path_info = self.find_app_path(app_name)

        if not path_info or not path_info[0]:
            self.speak(f"Could not find {app_name}")
            return False

        path, requires_admin = path_info

        try:
            if admin or requires_admin:
                if not self.is_admin():
                    self.speak(f"Attempting to open {app_name} as administrator")
                    return self.run_as_admin(path)

            if path.endswith(".exe"):
                os.startfile(path)
            else:
                subprocess.Popen(path, shell=True)

            self.speak(f"Opened {app_name}" + (" as administrator" if admin else ""))

            if app_name not in self.system_commands:
                self.learned_apps[app_name] = {
                    "path": path,
                    "requires_admin": requires_admin or admin,
                    "last_used": str(datetime.now())
                }
                self.save_memory()

            with open("activity.log", "a") as log:
                log.write(f"{datetime.now()}: Opened {app_name}\n")

            return True
        except Exception as e:
            self.speak(f"Failed to open {app_name}: {str(e)}")
            return False

    def process_command(self, command):
        command = command.lower()
        admin = " as admin" in command or "administrator" in command
        clean_command = command.replace(" as admin", "").replace("administrator", "").strip()

        if "add path" in command:
            self.add_app_path_interactively()
            return

        if "open routines" in clean_command or "routine manager" in clean_command:
            try:
                from routine_gui import RoutineManagerGUI
                self.speak("Opening routine manager.")
                gui = RoutineManagerGUI()
                gui.mainloop()
            except Exception as e:
                self.speak("Failed to open routine manager.")
                print(f"GUI error: {e}")
            return

        if "open" in clean_command:
            app_name = clean_command.replace("open", "").strip()
            self.open_app(app_name, admin)

        elif "close" in command:
            app_name = command.replace("close", "").strip()
            self.close_app(app_name)

        else:
            self.speak("Command not recognized")

    def run(self):
        self.speak("System controller ready")
        while True:
            cmd = self.listen()
            if cmd:
                print(f"USER: {cmd}")
                if any(word in cmd for word in ["exit", "quit", "close assistant", "stop"]):
                    self.speak("Goodbye!")
                    break
                self.process_command(cmd)

if __name__ == "__main__":
    controller = CompleteSystemController()
    if "--admin" in sys.argv:
        controller.run()
    elif controller.is_admin():
        controller.run()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv + ["--admin"]), None, 1)
