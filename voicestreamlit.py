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
import streamlit as st
import pythoncom

# Initialize COM library
pythoncom.CoInitialize()

# Initialize SAPI.SpVoice
try:
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
except Exception as e:
    st.write(f"Failed to initialize SAPI.SpVoice: {e}")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        st.write("Listening...")
        speaker.Speak("Listening...")
        try:
            audio = r.listen(source, timeout=5)
            st.write("Recognizing...")
            speaker.Speak("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            st.write(f"User said: {query}")
            speaker.Speak(f"You said: {query}")
            return query
        except sr.UnknownValueError:
            speaker.Speak("Sorry, I couldn't understand what you said.")
            return "Sorry, I couldn't understand what you said."
        except sr.RequestError as e:
            speaker.Speak("I'm having trouble connecting to the Google API.")
            st.write(f"Could not request results; {e}")
            return "I'm having trouble connecting to the Google API."
        except Exception as e:
            speaker.Speak("An error occurred.")
            st.write(f"Error: {e}")
            return f"Error: {e}"

# Your other functions...
def open_application(application_path):
    try:
        os.startfile(application_path)
    except Exception as e:
        speaker.Speak(f"Failed to open application: {e}")
        st.write(f"Failed to open application: {e}")


def open_camera():
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        if ret:
            cv2.imshow('Camera', frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
    cap.release()
    cv2.destroyAllWindows()

    
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
        time.sleep(5)
        pyautogui.press("enter")
        time.sleep(5)  # Wait for the chat to open

        # Click on the chat area to ensure the cursor is in the message input box
        # Update coordinates based on your screen resolution
        pyautogui.click(x=600, y=800)  # Example coordinates; adjust as needed
        
        # Add an additional small delay to ensure focus
        time.sleep(1)

        # Type the message and press Enter to send
        pyautogui.write(message)
        pyautogui.press("enter")
        time.sleep(2)
        
        print(f"Message sent to {contact_name}.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    st.title("Voice Command Assistant")

    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speaker.Speak("Good Morning, How Can I help You")
    elif hour >= 12 and hour < 18:
        speaker.Speak("Good Afternoon, How Can I help You")
    else:
        speaker.Speak("Good Evening, How Can I help You")

    if st.button('Give Command'):
        query = takeCommand()

        # Debugging output
        st.write(f"Received command: {query}")

        if "hi" in query.lower():
            speaker.Speak("Hello..How can I help You?")
            st.write("Responded to 'hi'")
        elif "open youtube" in query.lower():
            webbrowser.open("https://youtube.com")
            speaker.Speak("Opening YouTube")
            st.write("Opening YouTube")
        elif "open google" in query.lower():
            webbrowser.open("https://www.google.com")
            speaker.Speak("Opening Google")
            st.write("Opening Google")
        elif "tell me the time" in query.lower():
            current_time = datetime.datetime.now().strftime("%H:%M")
            speaker.Speak(f"The current time is {current_time}")
            st.write(f"The current time is {current_time}")
        elif "thank you" in query.lower():
            speaker.Speak("You're welcome!")
            st.write("Responded to 'thank you'")
        elif "open gmail" in query.lower():
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
            speaker.Speak("Opening Gmail, madam...")
        elif "open whatsapp" in query.lower():
            webbrowser.open("https://web.whatsapp.com/")
            speaker.Speak("Opening whatsapp, madam...")
        elif "open wikipedia" in query.lower():
            webbrowser.open("https://www.wikipedia.com")
            speaker.Speak("Opening wikipedia, madam...")
        elif "tell me the date" in query.lower():
            current_date = datetime.datetime.now().strftime("%Y-%m-%d")
            speaker.Speak(f"Today's date is {current_date}")
        elif "give system information" in query.lower():
            cpu_usage = psutil.cpu_percent(interval=1)
            ram_usage = psutil.virtual_memory().percent
            disk_usage = psutil.disk_usage('/').percent

            info_text = f"CPU Usage: {cpu_usage}%\nRAM Usage: {ram_usage}%\nDisk Usage: {disk_usage}%"
            speaker.Speak(info_text)
        elif "send whatsapp message" in query.lower():
            speaker.Speak("Whom do you want to send a message to?")
            contact_name = takeCommand()
            speaker.Speak("What message do you want to send?")
            message1 = takeCommand()

            send_whatsapp_message(contact_name, message1)
        elif "wikipedia" in query.lower():
            speaker.Speak("You want to search Wikipedia. Please tell me what you would like to search for.")
            search_query = takeCommand()
            st.write(f"Searching Wikipedia for: {search_query}")
            try:
                results = wikipedia.summary(search_query, sentences=2)
        
                speaker.Speak("According to Wikipedia")
                speaker.Speak(results)
                st.write(results)
            except wikipedia.exceptions.DisambiguationError as e:
                speaker.Speak("The query was ambiguous. Here are some suggestions.")
                st.write(f"Disambiguation Error: {e}")
            except wikipedia.exceptions.PageError:
                speaker.Speak("No results found on Wikipedia for the given query.")
                st.write("Page not found on Wikipedia.")
            except Exception as e:
                speaker.Speak("An error occurred while searching Wikipedia.")
                st.write(f"Error: {e}")

        elif "open music" in query.lower():
            musicPath = r"C:\Users\nooth\Downloads\Music.mp3"
            os.system(f'start explorer "{musicPath}"')
            speaker.Speak("have fun with music baby")
        elif "open camera" in query.lower():
            speaker.Speak("Look at you beautiful!!")
            open_camera()
            #speaker.Speak("Lokk at you beautiful")
        elif "open notepad" in query.lower():
            open_application("notepad.exe")
            st.write("opening notepad")
        elif "open calculator" in query.lower():
            open_application("calc.exe")
            st.write("opening calci")
        elif "open file explorer" in query.lower():
            open_application("explorer.exe")
        elif "thankyou" in query.lower():
            speaker.Speak("Your Welcome,babe!!")
        elif "open word" in query.lower():
            open_application("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")
        elif "open excel" in query.lower():
            open_application("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
        elif "open powerpoint" in query.lower():
            open_application("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk")

        else:
            speaker.Speak("Sorry, I didn't understand that.")
            st.write("Unknown command")

    # Uninitialize COM library
    pythoncom.CoUninitialize()