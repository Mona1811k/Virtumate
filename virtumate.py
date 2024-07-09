import datetime
import os
import subprocess
import time
import webbrowser
import dateparser
import schedule
import cv2
import psutil
import speech_recognition as sr
import win32com.client
import pyautogui
import wikipedia

speaker = win32com.client.Dispatch("SAPI.SpVoice")
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        print("Listening...")
        try:
            audio = r.listen(source, timeout=5)
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            return "Sorry, I couldn't understand what you said"
        except sr.RequestError:
            return "I'm having trouble connecting to the Google API"


def set_reminder():
    speaker.Speak("Sure! When do you want me to remind you?")
    print("Sure! When do you want me to remind you?")

    reminder_time_phrase = takeCommand().lower()

    # Parse the time phrase into a datetime object
    reminder_datetime = dateparser.parse(reminder_time_phrase)

    if reminder_datetime:
        # Extract hour and minute
        hour = reminder_datetime.hour
        minute = reminder_datetime.minute

        # Schedule the reminder
        schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(remind_user,
                                                               "Reminder: It's time to...")  # Adjust the reminder message
        print(f"Reminder scheduled for {hour:02d}:{minute:02d}.")
        speaker.Speak(f"Reminder scheduled for {hour:02d}:{minute:02d}.")
    else:
        print("Sorry, I couldn't understand the time. Please try again.")
        speaker.Speak("Sorry, I couldn't understand the time. Please try again.")
def remind_user(reminder_text, delay):
    time_format = "%Y-%m-%d %H:%M:%S"
    now = datetime.datetime.now()
    reminder_time = now + datetime.timedelta(seconds=delay)
    print(f"Reminder set for: {reminder_time.strftime(time_format)}")
    while datetime.datetime.now() < reminder_time:
        pass
    speaker.Speak(reminder_text)

def send_whatsapp_message(contact_name, message):
    try:
        # Open the Start menu to switch to WhatsApp
        pyautogui.hotkey("win")
        time.sleep(1)  # Wait for the Start menu to open

        # Type "WhatsApp" and press Enter to open it
        pyautogui.write("WhatsApp")
        pyautogui.press("enter")
        time.sleep(5)  # Wait for WhatsApp to open

        # Type the contact name in the search bar and press Enter
        pyautogui.write(contact_name)
        pyautogui.press("enter")
        time.sleep(2)  # Wait for the contact to be selected

        # Type the message and press Enter to send
        pyautogui.write(message)
        pyautogui.press("enter")

        speaker.Speak(f"Message sent to {contact_name}.")
        print(f"Message sent to {contact_name}.")
    except Exception as e:
        print(f"An error occurred: {e}")


def open_application(application_name):
    try:
        # Try to open the application using the default system command
        subprocess.Popen(application_name, shell=True)
    except Exception as e:
        print(f"Error: {e}")


def open_camera():
    # Open the default camera (camera index 0)
    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        print("Error: Could not open camera.")
        return

    while True:
        # Read a frame from the camera
        ret, frame = cap.read()

        # Display the frame
        cv2.imshow("Camera", frame)

        # Break the loop if 'q' key is pressed
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Release the camera
    cap.release()

    # Close all OpenCV windows
    cv2.destroyAllWindows()


if _name_ == '_main_':
    try:
        hour=int(datetime.datetime.now().hour)
        if hour>=0 and hour<12:
            speaker.Speak("Good Morning,How Can I help You")
        elif hour>=12 and hour<18:
            speaker.Speak("Good Afternoon,How Can I help You")
        else:
            speaker.Speak("Good Evening,How Can I help You")


        while True:
            query = takeCommand()
            if "hi " in query.lower():
                speaker.Speak("Hello..How can I help You?")
            if "open youtube" in query.lower():
                webbrowser.open("https://youtube.com")
                speaker.Speak("Opening YouTube")
            if "open gmail" in query.lower():
                webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
                speaker.Speak("Opening Gmail, madam...")
            if "open whatsapp" in query.lower():
                webbrowser.open("https://web.whatsapp.com/")
                speaker.Speak("Opening whatsapp, madam...")
            if "open wikipedia" in query.lower():
                webbrowser.open("https://www.wikipedia.com")
                speaker.Speak("Opening wikipedia, madam...")
            if "open google" in query.lower():
                webbrowser.open("https://www.google.com")
                speaker.Speak("Opening Google, madam...")
            if "set a reminder" in query:
                set_reminder()
            if "tell me the time" in query.lower():
                current_time = datetime.datetime.now().strftime("%H:%M")
                speaker.Speak(f"The current time is {current_time}")
            if "tell me the date" in query.lower():
                current_date = datetime.datetime.now().strftime("%Y-%m-%d")
                speaker.Speak(f"Today's date is {current_date}")
            if "give system information" in query.lower():
                cpu_usage = psutil.cpu_percent(interval=1)
                ram_usage = psutil.virtual_memory().percent
                disk_usage = psutil.disk_usage('/').percent

                info_text = f"CPU Usage: {cpu_usage}%\nRAM Usage: {ram_usage}%\nDisk Usage: {disk_usage}%"
                speaker.Speak(info_text)
            if "send whatsapp message" in query.lower():
                speaker.Speak("Whom do you want to send a message to?")
                contact_name = takeCommand()
                speaker.Speak("What message do you want to send?")
                message = takeCommand()

                send_whatsapp_message(contact_name, message)
            if "wikipedia" in query.lower():
                speaker.Speak("Searching wikipedia...")
                query=query.replace("wikipedia", "")
                results=wikipedia.summary(query, sentences=2)
                speaker.Speak("according to wikipedia")
                speaker.Speak(results)

            if "send email" in query.lower():
                speaker.Speak("Whom do you want to send an email to?")
                recipient = takeCommand()
                speaker.Speak("What is the subject of the email?")
                subject = takeCommand()
                speaker.Speak("What is the content of the email?")
                body = takeCommand()
                send_email_jarvis(recipient, subject, body)

            if "open music" in query.lower():
                musicPath = r"C:\Users\nooth\Downloads\Music.mp3"
                os.system(f'start explorer "{musicPath}"')
            if "open camera" in query.lower():
                open_camera()
            if "open notepad" in query.lower():
                open_application("notepad.exe")
            if "open calculator" in query.lower():
                open_application("calc.exe")
            if "open file explorer" in query.lower():
                open_application("explorer.exe")
            if "thankyou" in query.lower():
                speaker.Speak("Your Welcome!!")
            if "thankyou jarvis" in query.lower():
                speaker.Speak("Your Welcome!!")
            if "open word" in query.lower():
                open_application("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")
            if "open excel" in query.lower():
                open_application("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
            if "open powerpoint" in query.lower():
                open_application("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk")
            schedule.run_pending()
            time.sleep(1)
    except Exception as e:
        print(f"An error occurred: {e}")