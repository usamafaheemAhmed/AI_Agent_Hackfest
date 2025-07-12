#!/usr/bin/env python3
"""
AI-Powered PC Automation Agent
A hackathon project that uses AI to understand and execute computer automation tasks.
Supports both voice and text input with various system automation capabilities.
"""

import os
import sys
import json
import time
import threading
import webbrowser
import subprocess
import pyautogui
import pyttsx3
import speech_recognition as sr
from openai import OpenAI
from typing import Dict, List, Optional
import keyboard
import tkinter as tk
from tkinter import ttk, scrolledtext
import queue

# Configure pyautogui safety
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

class AIAutomationAgent:
    def __init__(self):
        """Initialize the AI automation agent with all necessary components."""
        self.setup_openai()
        self.setup_speech_recognition()
        self.setup_tts()
        self.setup_gui()
        self.command_queue = queue.Queue()
        self.is_listening = False
        self.is_running = True
        
        # Define system commands and their implementations
        self.system_commands = {
            'open_chrome': self.open_chrome,
            'open_vscode': self.open_vscode,
            'open_notepad': self.open_notepad,
            'open_calculator': self.open_calculator,
            'google_search': self.google_search,
            'youtube_search': self.youtube_search,
            'type_text': self.type_text,
            'play_pause_media': self.play_pause_media,
            'next_tab': self.next_tab,
            'previous_tab': self.previous_tab,
            'close_tab': self.close_tab,
            'new_tab': self.new_tab,
            'minimize_window': self.minimize_window,
            'maximize_window': self.maximize_window,
            'take_screenshot': self.take_screenshot,
            'lock_screen': self.lock_screen,
            'open_file_explorer': self.open_file_explorer,
            'volume_up': self.volume_up,
            'volume_down': self.volume_down,
            'mute_unmute': self.mute_unmute
        }

    def setup_openai(self):
        """Set up OpenAI client with API key."""
        try:
            # Try to get API key from environment variable
            api_key = os.getenv('OPENAI_API_KEY')
            if not api_key:
                # If not found, prompt user
                print("OpenAI API key not found in environment variables.")
                print("Please set OPENAI_API_KEY environment variable or enter it now:")
                api_key = input("Enter your OpenAI API key: ").strip()
                if not api_key:
                    raise ValueError("OpenAI API key is required")
            
            self.openai_client = OpenAI(api_key=api_key)
            self.log_message("✅ OpenAI client initialized successfully")
        except Exception as e:
            self.log_message(f"❌ Error setting up OpenAI: {str(e)}")
            sys.exit(1)

    def setup_speech_recognition(self):
        """Set up speech recognition components."""
        try:
            self.recognizer = sr.Recognizer()
            self.microphone = sr.Microphone()
            
            # Adjust for ambient noise
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
            
            self.log_message("✅ Speech recognition initialized successfully")
        except Exception as e:
            self.log_message(f"❌ Error setting up speech recognition: {str(e)}")

    def setup_tts(self):
        """Set up text-to-speech engine."""
        try:
            self.tts_engine = pyttsx3.init()
            self.tts_engine.setProperty('rate', 200)  # Speed of speech
            self.tts_engine.setProperty('volume', 0.8)  # Volume level
            self.log_message("✅ Text-to-speech engine initialized successfully")
        except Exception as e:
            self.log_message(f"❌ Error setting up TTS: {str(e)}")

    def setup_gui(self):
        """Set up the graphical user interface."""
        self.root = tk.Tk()
        self.root.title("AI PC Automation Agent")
        self.root.geometry("800x600")
        self.root.configure(bg='#2b2b2b')
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="🤖 AI PC Automation Agent", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Text input frame
        input_frame = ttk.LabelFrame(main_frame, text="Command Input", padding="10")
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Text input
        self.text_input = ttk.Entry(input_frame, width=60, font=("Arial", 12))
        self.text_input.grid(row=0, column=0, padx=(0, 10))
        self.text_input.bind('<Return>', self.on_text_submit)
        
        # Buttons frame
        buttons_frame = ttk.Frame(input_frame)
        buttons_frame.grid(row=0, column=1)
        
        # Submit button
        submit_btn = ttk.Button(buttons_frame, text="Execute", command=self.on_text_submit)
        submit_btn.grid(row=0, column=0, padx=(0, 5))
        
        # Voice button
        self.voice_btn = ttk.Button(buttons_frame, text="🎤 Voice", command=self.toggle_voice_listening)
        self.voice_btn.grid(row=0, column=1)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Activity Log", padding="10")
        log_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, width=80, 
                                                 font=("Consolas", 10), bg='#1e1e1e', fg='#ffffff')
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Sample commands frame
        samples_frame = ttk.LabelFrame(main_frame, text="Sample Commands", padding="10")
        samples_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        sample_commands = [
            "Open Chrome and search for AI hackathon",
            "Open VS Code",
            "Type 'Hello World'",
            "Play or pause media",
            "Take a screenshot",
            "Open YouTube and search for Python tutorials"
        ]
        
        samples_text = "\n".join(f"• {cmd}" for cmd in sample_commands)
        samples_label = ttk.Label(samples_frame, text=samples_text, justify=tk.LEFT)
        samples_label.grid(row=0, column=0, sticky=tk.W)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Start command processing thread
        self.command_thread = threading.Thread(target=self.process_commands, daemon=True)
        self.command_thread.start()

    def log_message(self, message: str):
        """Add a message to the log with timestamp."""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
        else:
            print(log_entry.strip())

    def speak(self, text: str):
        """Convert text to speech."""
        try:
            self.tts_engine.say(text)
            self.tts_engine.runAndWait()
        except Exception as e:
            self.log_message(f"❌ TTS Error: {str(e)}")

    def listen_for_voice(self):
        """Listen for voice input and convert to text."""
        try:
            with self.microphone as source:
                self.log_message("🎤 Listening for voice input...")
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
            
            self.log_message("🔄 Processing voice input...")
            text = self.recognizer.recognize_google(audio)
            self.log_message(f"📝 Voice input recognized: {text}")
            return text
        except sr.WaitTimeoutError:
            self.log_message("⏰ Voice input timeout")
            return None
        except sr.UnknownValueError:
            self.log_message("❌ Could not understand voice input")
            return None
        except sr.RequestError as e:
            self.log_message(f"❌ Speech recognition error: {str(e)}")
            return None

    def parse_command_with_ai(self, user_input: str) -> Dict:
        """Use OpenAI to parse and understand the user's command."""
        try:
            system_prompt = """You are an AI assistant that helps parse user commands for PC automation.
            
            Available commands:
            - open_chrome: Open Google Chrome browser
            - open_vscode: Open Visual Studio Code
            - open_notepad: Open Notepad
            - open_calculator: Open Calculator
            - google_search: Search on Google (requires 'query' parameter)
            - youtube_search: Search on YouTube (requires 'query' parameter)
            - type_text: Type text automatically (requires 'text' parameter)
            - play_pause_media: Play or pause media
            - next_tab: Switch to next browser tab
            - previous_tab: Switch to previous browser tab
            - close_tab: Close current browser tab
            - new_tab: Open new browser tab
            - minimize_window: Minimize current window
            - maximize_window: Maximize current window
            - take_screenshot: Take a screenshot
            - lock_screen: Lock the computer screen
            - open_file_explorer: Open file explorer
            - volume_up: Increase volume
            - volume_down: Decrease volume
            - mute_unmute: Mute or unmute audio
            
            Parse the user's command and return a JSON response with:
            - "command": the appropriate command from the list above
            - "parameters": any required parameters (like query for search, text for typing)
            - "description": a brief description of what will be executed
            
            If the command is unclear or not supported, return:
            - "command": "unknown"
            - "error": explanation of the issue
            """
            
            response = self.openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_input}
                ],
                max_tokens=200,
                temperature=0.3
            )
            
            ai_response = response.choices[0].message.content.strip()
            self.log_message(f"🤖 AI Response: {ai_response}")
            
            # Parse JSON response
            try:
                parsed_command = json.loads(ai_response)
                return parsed_command
            except json.JSONDecodeError:
                # If JSON parsing fails, try to extract key information
                return {
                    "command": "unknown",
                    "error": "Failed to parse AI response as JSON"
                }
                
        except Exception as e:
            self.log_message(f"❌ Error calling OpenAI API: {str(e)}")
            return {
                "command": "unknown",
                "error": f"API Error: {str(e)}"
            }

    def execute_command(self, command_data: Dict):
        """Execute the parsed command."""
        try:
            command = command_data.get('command', '')
            parameters = command_data.get('parameters', {})
            
            if command == 'unknown':
                error_msg = command_data.get('error', 'Unknown command')
                self.log_message(f"❌ {error_msg}")
                self.speak("I'm sorry, I didn't understand that command.")
                return
            
            if command not in self.system_commands:
                self.log_message(f"❌ Command '{command}' not implemented")
                self.speak("This command is not yet implemented.")
                return
            
            # Execute the command
            self.log_message(f"⚡ Executing: {command}")
            if command_data.get('description'):
                self.speak(f"Executing: {command_data['description']}")
            
            # Call the appropriate method
            if parameters:
                self.system_commands[command](**parameters)
            else:
                self.system_commands[command]()
            
            self.log_message(f"✅ Command '{command}' executed successfully")
            
        except Exception as e:
            self.log_message(f"❌ Error executing command: {str(e)}")
            self.speak("Sorry, there was an error executing that command.")

    # System command implementations
    def open_chrome(self):
        """Open Google Chrome browser."""
        try:
            if sys.platform == "win32":
                os.startfile("chrome")
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", "-a", "Google Chrome"])
            else:  # Linux
                subprocess.run(["google-chrome"])
        except Exception as e:
            self.log_message(f"❌ Error opening Chrome: {str(e)}")

    def open_vscode(self):
        """Open Visual Studio Code."""
        try:
            if sys.platform == "win32":
                subprocess.run(["code"])
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", "-a", "Visual Studio Code"])
            else:  # Linux
                subprocess.run(["code"])
        except Exception as e:
            self.log_message(f"❌ Error opening VS Code: {str(e)}")

    def open_notepad(self):
        """Open Notepad."""
        try:
            if sys.platform == "win32":
                subprocess.run(["notepad"])
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", "-a", "TextEdit"])
            else:  # Linux
                subprocess.run(["gedit"])
        except Exception as e:
            self.log_message(f"❌ Error opening Notepad: {str(e)}")

    def open_calculator(self):
        """Open Calculator."""
        try:
            if sys.platform == "win32":
                subprocess.run(["calc"])
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", "-a", "Calculator"])
            else:  # Linux
                subprocess.run(["gnome-calculator"])
        except Exception as e:
            self.log_message(f"❌ Error opening Calculator: {str(e)}")

    def google_search(self, query: str):
        """Perform a Google search."""
        try:
            search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            webbrowser.open(search_url)
        except Exception as e:
            self.log_message(f"❌ Error performing Google search: {str(e)}")

    def youtube_search(self, query: str):
        """Perform a YouTube search."""
        try:
            search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
            webbrowser.open(search_url)
        except Exception as e:
            self.log_message(f"❌ Error performing YouTube search: {str(e)}")

    def type_text(self, text: str):
        """Type text automatically."""
        try:
            time.sleep(2)  # Give user time to click where they want to type
            pyautogui.typewrite(text)
        except Exception as e:
            self.log_message(f"❌ Error typing text: {str(e)}")

    def play_pause_media(self):
        """Play or pause media."""
        try:
            pyautogui.press('playpause')
        except Exception as e:
            self.log_message(f"❌ Error controlling media: {str(e)}")

    def next_tab(self):
        """Switch to next browser tab."""
        try:
            pyautogui.hotkey('ctrl', 'tab')
        except Exception as e:
            self.log_message(f"❌ Error switching to next tab: {str(e)}")

    def previous_tab(self):
        """Switch to previous browser tab."""
        try:
            pyautogui.hotkey('ctrl', 'shift', 'tab')
        except Exception as e:
            self.log_message(f"❌ Error switching to previous tab: {str(e)}")

    def close_tab(self):
        """Close current browser tab."""
        try:
            pyautogui.hotkey('ctrl', 'w')
        except Exception as e:
            self.log_message(f"❌ Error closing tab: {str(e)}")

    def new_tab(self):
        """Open new browser tab."""
        try:
            pyautogui.hotkey('ctrl', 't')
        except Exception as e:
            self.log_message(f"❌ Error opening new tab: {str(e)}")

    def minimize_window(self):
        """Minimize current window."""
        try:
            pyautogui.hotkey('win', 'down')
        except Exception as e:
            self.log_message(f"❌ Error minimizing window: {str(e)}")

    def maximize_window(self):
        """Maximize current window."""
        try:
            pyautogui.hotkey('win', 'up')
        except Exception as e:
            self.log_message(f"❌ Error maximizing window: {str(e)}")

    def take_screenshot(self):
        """Take a screenshot."""
        try:
            screenshot = pyautogui.screenshot()
            filename = f"screenshot_{int(time.time())}.png"
            screenshot.save(filename)
            self.log_message(f"📸 Screenshot saved as: {filename}")
        except Exception as e:
            self.log_message(f"❌ Error taking screenshot: {str(e)}")

    def lock_screen(self):
        """Lock the computer screen."""
        try:
            if sys.platform == "win32":
                pyautogui.hotkey('win', 'l')
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["osascript", "-e", 'tell application "System Events" to keystroke "q" using {command down, control down}'])
            else:  # Linux
                subprocess.run(["xdg-screensaver", "lock"])
        except Exception as e:
            self.log_message(f"❌ Error locking screen: {str(e)}")

    def open_file_explorer(self):
        """Open file explorer."""
        try:
            if sys.platform == "win32":
                pyautogui.hotkey('win', 'e')
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", "-a", "Finder"])
            else:  # Linux
                subprocess.run(["nautilus"])
        except Exception as e:
            self.log_message(f"❌ Error opening file explorer: {str(e)}")

    def volume_up(self):
        """Increase volume."""
        try:
            pyautogui.press('volumeup')
        except Exception as e:
            self.log_message(f"❌ Error increasing volume: {str(e)}")

    def volume_down(self):
        """Decrease volume."""
        try:
            pyautogui.press('volumedown')
        except Exception as e:
            self.log_message(f"❌ Error decreasing volume: {str(e)}")

    def mute_unmute(self):
        """Mute or unmute audio."""
        try:
            pyautogui.press('volumemute')
        except Exception as e:
            self.log_message(f"❌ Error muting/unmuting: {str(e)}")

    def on_text_submit(self, event=None):
        """Handle text input submission."""
        user_input = self.text_input.get().strip()
        if user_input:
            self.log_message(f"📝 Text input: {user_input}")
            self.command_queue.put(user_input)
            self.text_input.delete(0, tk.END)

    def toggle_voice_listening(self):
        """Toggle voice listening on/off."""
        if not self.is_listening:
            self.is_listening = True
            self.voice_btn.config(text="🔴 Stop")
            threading.Thread(target=self.voice_listening_thread, daemon=True).start()
        else:
            self.is_listening = False
            self.voice_btn.config(text="🎤 Voice")

    def voice_listening_thread(self):
        """Thread for continuous voice listening."""
        while self.is_listening:
            try:
                voice_input = self.listen_for_voice()
                if voice_input:
                    self.command_queue.put(voice_input)
                time.sleep(0.1)
            except Exception as e:
                self.log_message(f"❌ Voice listening error: {str(e)}")
                break
        
        self.is_listening = False
        self.voice_btn.config(text="🎤 Voice")

    def process_commands(self):
        """Process commands from the queue."""
        while self.is_running:
            try:
                # Wait for command with timeout
                user_input = self.command_queue.get(timeout=0.1)
                
                # Parse command with AI
                command_data = self.parse_command_with_ai(user_input)
                
                # Execute command
                self.execute_command(command_data)
                
            except queue.Empty:
                continue
            except Exception as e:
                self.log_message(f"❌ Error processing command: {str(e)}")

    def run(self):
        """Start the main application loop."""
        self.log_message("🚀 AI PC Automation Agent started!")
        self.log_message("💡 Type commands or use voice input to control your PC")
        self.log_message("📋 Check the sample commands below for ideas")
        
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.log_message("👋 Shutting down...")
        finally:
            self.is_running = False

def main():
    """Main entry point of the application."""
    print("🤖 AI PC Automation Agent")
    print("=" * 50)
    print("Setting up the application...")
    
    # Create and run the agent
    agent = AIAutomationAgent()
    agent.run()

if __name__ == "__main__":
    main()