# Virtumate
This project implements a simple voice assistant using various Python libraries to perform tasks such as setting reminders, sending WhatsApp messages, opening applications, and providing system information. The assistant listens for voice commands and executes the corresponding actions.

Libraries Used:
1)datetime: Provides classes for manipulating dates and times.
2)os: Provides functions for interacting with the operating system.
3)subprocess: Allows for spawning new processes and connecting to their input/output/error pipes.
4)time: Provides time-related functions.
5)webbrowser: Allows for opening URLs in a web browser.
6)dateparser: Parses dates from natural language strings.
7)schedule: Allows for scheduling jobs at specific intervals.
8)cv2: OpenCV library for computer vision tasks.
9)psutil: Provides an interface for retrieving information on running processes and system utilization.
10)speech_recognition: Recognizes speech and converts it to text.
11)win32com.client: Provides access to COM objects, used here for text-to-speech functionality.
12)pyautogui: Provides functions for controlling the mouse and keyboard.
13)wikipedia: Allows for searching and retrieving summaries from Wikipedia.

Functionalities
1. Voice Command Recognition
The takeCommand function listens for voice input using the microphone and recognizes speech using Google's speech recognition API.
2. Setting Reminders
The set_reminder function allows the user to set reminders by parsing natural language time phrases and scheduling reminders.
3.The open_application function opens specified applications using the subprocess library.
4.The open_application function opens specified applications using the subprocess library.
5.The main functionality includes various voice commands to open websites, set reminders, tell the time and date, provide system information, send WhatsApp messages, search Wikipedia, send emails, play music, open applications, and thank the user.
