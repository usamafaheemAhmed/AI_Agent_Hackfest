#!/usr/bin/env python3
"""
AI-Powered PC Automation Agent
A hackathon project that uses AI to understand and execute computer automation tasks.
Supports both voice and text input with various system automation capabilities.
Now using Google Gemini API instead of OpenAI.
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
import google.generativeai as genai
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
        # Initialize basic attributes first
        self.command_queue = queue.Queue()
        self.is_listening = False
        self.is_running = True
        
        # Setup components
        self.setup_gemini()
        self.setup_speech_recognition()
        self.setup_tts()
        self.setup_gui()
        
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

    def setup_gemini(self):
        """Set up Google Gemini API client with API key."""
        try:
            # Try to get API key from environment variable
            api_key = os.getenv('GOOGLE_API_KEY')
            if not api_key:
                # If not found, prompt user
                print("Google API key not found in environment variables.")
                print("Please set GOOGLE_API_KEY environment variable or enter it now:")
                api_key = input("Enter your Google API key: ").strip()
                if not api_key:
                    raise ValueError("Google API key is required")
            
            # Configure the API
            genai.configure(api_key=api_key)
            
            # Initialize the model
            self.model = genai.GenerativeModel('gemini-1.5-flash')
            
            self.log_message("✅ Google Gemini API initialized successfully")
        except Exception as e:
            self.log_message(f"❌ Error setting up Gemini API: {str(e)}")
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
        self.root.title("AI PC Automation Agent (Gemini)")
        self.root.geometry("800x600")
        self.root.configure(bg='#2b2b2b')
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="🤖 AI PC Automation Agent (Gemini)", 
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
        
        # Start command processing thread after all initialization is complete
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
            # Run TTS in a separate thread to avoid blocking
            def tts_thread():
                self.tts_engine.say(text)
                self.tts_engine.runAndWait()
            
            threading.Thread(target=tts_thread, daemon=True).start()
            self.log_message(f"🔊 Speaking: {text}")
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
        """Use Google Gemini to parse and understand the user's command."""
        try:
            self.log_message(f"🤖 Sending to Gemini: {user_input}")
            
            prompt = f"""You are an AI assistant that helps parse user commands for PC automation.
            
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
            
            Parse the user's command and return ONLY a JSON response with:
            - "command": the appropriate command from the list above, or "unknown" if no match
            - "parameters": any required parameters (like query for search, text for typing)
            - "description": a brief description of what will be executed

            If no command matches, return: {{"command": "unknown", "parameters": {{}}, "description": "Command not recognized"}}            
            Examples:
            Input: "open vs code" -> {{"command": "open_vscode", "parameters": {{}}, "description": "Opening Visual Studio Code"}}
            Input: "search google for python" -> {{"command": "google_search", "parameters": {{"query": "python"}}, "description": "Searching Google for python"}}
            Input: "type hello world" -> {{"command": "type_text", "parameters": {{"text": "hello world"}}, "description": "Typing hello world"}}
            Input: "hello world" -> {{"command": "type_text", "parameters": {{"text": "hello world"}}, "description": "Typing hello world"}}
            Input: "write something" -> {{"command": "type_text", "parameters": {{"text": "something"}}, "description": "Typing something"}}
            
            Special rules:
            - If user provides plain text without a clear command verb, interpret it as a type_text command
            - Commands like "hello world", "write this", or any plain text should use type_text
            - Always try to match to an available command rather than returning unknown
            
            User command: "{user_input}"
            
            Return ONLY the JSON, no other text."""
            
            # Generate response using Gemini
            response = self.model.generate_content(prompt)
            ai_response = response.text.strip()
            
            self.log_message(f"🤖 Gemini Response: {ai_response}")
            
            # Clean up the response (remove code blocks if present)
            if ai_response.startswith("```"):
                lines = ai_response.split('\n')
                # Find the JSON part
                json_start = -1
                json_end = -1
                for i, line in enumerate(lines):
                    if line.strip().startswith('{'):
                        json_start = i
                    if line.strip().endswith('}') and json_start != -1:
                        json_end = i
                        break
                
                if json_start != -1 and json_end != -1:
                    ai_response = '\n'.join(lines[json_start:json_end+1])
            
            # Remove any remaining markdown formatting
            ai_response = ai_response.replace('```json', '').replace('```', '').strip()
            
            # Parse JSON response
            try:
                parsed_command = json.loads(ai_response)
                # Handle null command
                if parsed_command.get('command') is None:
                    parsed_command['command'] = 'unknown'
                return parsed_command
            except json.JSONDecodeError as e:
                self.log_message(f"❌ JSON Parse Error: {str(e)}")
                # Try to extract JSON from the response
                try:
                    # Look for JSON-like structure in the response
                    start_idx = ai_response.find('{')
                    end_idx = ai_response.rfind('}') + 1
                    if start_idx != -1 and end_idx != -1:
                        json_str = ai_response[start_idx:end_idx]
                        parsed_command = json.loads(json_str)
                        return parsed_command
                except:
                    pass
                
                # Fallback to simple parsing
                return self.fallback_parse_command(user_input)
                
        except Exception as e:
            self.log_message(f"❌ Error calling Gemini API: {str(e)}")
            return self.fallback_parse_command(user_input)

    def fallback_parse_command(self, user_input: str) -> Dict:
        """Fallback command parsing without AI."""
        user_lower = user_input.lower()
        
        # Simple keyword matching
        if "vs code" in user_lower or "vscode" in user_lower:
            return {"command": "open_vscode", "parameters": {}, "description": "Opening Visual Studio Code"}
        elif "chrome" in user_lower:
            return {"command": "open_chrome", "parameters": {}, "description": "Opening Google Chrome"}
        elif "notepad" in user_lower:
            return {"command": "open_notepad", "parameters": {}, "description": "Opening Notepad"}
        elif "calculator" in user_lower:
            return {"command": "open_calculator", "parameters": {}, "description": "Opening Calculator"}
        elif "google" in user_lower and "search" in user_lower:
            # Extract search query
            words = user_input.split()
            query_words = []
            found_search = False
            for word in words:
                if found_search:
                    query_words.append(word)
                elif word.lower() in ["search", "for", "google"]:
                    found_search = True
            query = " ".join(query_words) if query_words else "AI"
            return {"command": "google_search", "parameters": {"query": query}, "description": f"Searching Google for {query}"}
        elif "youtube" in user_lower and "search" in user_lower:
            # Extract search query
            words = user_input.split()
            query_words = []
            found_search = False
            for word in words:
                if found_search:
                    query_words.append(word)
                elif word.lower() in ["search", "for", "youtube"]:
                    found_search = True
            query = " ".join(query_words) if query_words else "music"
            return {"command": "youtube_search", "parameters": {"query": query}, "description": f"Searching YouTube for {query}"}
        elif "type" in user_lower:
            # Extract text to type
            words = user_input.split()
            text_words = []
            found_type = False
            for word in words:
                if found_type:
                    text_words.append(word)
                elif word.lower() == "type":
                    found_type = True
            text = " ".join(text_words) if text_words else "Hello World"
            return {"command": "type_text", "parameters": {"text": text}, "description": f"Typing {text}"}
        elif "play" in user_lower or "pause" in user_lower:
            return {"command": "play_pause_media", "parameters": {}, "description": "Playing or pausing media"}
        elif "screenshot" in user_lower:
            return {"command": "take_screenshot", "parameters": {}, "description": "Taking a screenshot"}
        elif "volume up" in user_lower or "increase volume" in user_lower:
            return {"command": "volume_up", "parameters": {}, "description": "Increasing volume"}
        elif "volume down" in user_lower or "decrease volume" in user_lower:
            return {"command": "volume_down", "parameters": {}, "description": "Decreasing volume"}
        elif "mute" in user_lower or "unmute" in user_lower:
            return {"command": "mute_unmute", "parameters": {}, "description": "Muting or unmuting audio"}
        else:
            return {"command": "unknown", "error": "Command not recognized"}

    def execute_command(self, command_data: Dict):
        """Execute the parsed command."""
        try:
            command = command_data.get('command', '')
            parameters = command_data.get('parameters', {})
            description = command_data.get('description', '')
            
            self.log_message(f"⚡ Executing command: {command}")
            
            if command == 'unknown':
                error_msg = command_data.get('error', 'Unknown command')
                self.log_message(f"❌ {error_msg}")
                self.speak("I'm sorry, I didn't understand that command.")
                return
            
            if command not in self.system_commands:
                self.log_message(f"❌ Command '{command}' not implemented")
                self.speak("This command is not yet implemented.")
                return
            
            # Speak what we're about to do
            if description:
                self.speak(description)
            else:
                self.speak(f"Executing {command}")
            
            # Execute the command
            if parameters:
                self.system_commands[command](**parameters)
            else:
                self.system_commands[command]()
            
            self.log_message(f"✅ Command '{command}' executed successfully")
            self.speak("Command completed successfully.")
            
        except Exception as e:
            error_msg = f"Error executing command: {str(e)}"
            self.log_message(f"❌ {error_msg}")
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
                # Try multiple common VS Code installation paths
                paths = [
                    "code",  # If in PATH
                    r"C:\Users\{}\AppData\Local\Programs\Microsoft VS Code\Code.exe".format(os.getenv('USERNAME')),
                    r"C:\Program Files\Microsoft VS Code\Code.exe",
                    r"C:\Program Files (x86)\Microsoft VS Code\Code.exe"
                ]
                
                success = False
                for path in paths:
                    try:
                        if path == "code":
                            subprocess.run([path], check=True)
                        else:
                            if os.path.exists(path):
                                subprocess.run([path], check=True)
                        success = True
                        break
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        continue
                
                if not success:
                    # Try opening through Windows start menu
                    subprocess.run(["start", "code"], shell=True)
                    
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", "-a", "Visual Studio Code"])
            else:  # Linux
                subprocess.run(["code"])
                
        except Exception as e:
            self.log_message(f"❌ Error opening VS Code: {str(e)}")
            # Try alternative method
            try:
                if sys.platform == "win32":
                    os.system("start code")
                else:
                    os.system("code")
            except Exception as e2:
                self.log_message(f"❌ Alternative VS Code open failed: {str(e2)}")

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
        self.log_message("🚀 AI PC Automation Agent started with Google Gemini!")
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
    print("🤖 AI PC Automation Agent (Google Gemini)")
    print("=" * 50)
    print("Setting up the application...")
    
    # Create and run the agent
    agent = AIAutomationAgent()
    agent.run()

if __name__ == "__main__":
    main()