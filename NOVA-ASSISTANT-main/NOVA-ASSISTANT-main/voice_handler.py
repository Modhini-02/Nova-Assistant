import speech_recognition as sr
import pyttsx3
import threading
import time

engine = None
engine_lock = threading.Lock()

def initialize_engine():
    global engine
    if engine is None:
        engine = pyttsx3.init()

def speak(text):
    global engine
    with engine_lock:
        if engine is None:
            initialize_engine()
        print(f"Nova: {text}")
        engine.say(text)
        engine.runAndWait()
        time.sleep(0.5)

def listen():
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 3000
    recognizer.pause_threshold = 1.0
    recognizer.dynamic_energy_threshold = True

    with sr.Microphone() as source:
        print("Listening...")
        try:
            audio = recognizer.listen(source, timeout=7, phrase_time_limit=10)
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text.lower()
        except sr.WaitTimeoutError:
            return "Timeout: No speech detected."
        except sr.UnknownValueError:
            return "Sorry, I didn't catch that."
        except sr.RequestError as e:
            return f"Speech service error: {e}"
        except Exception as e:
            return f"Error: {str(e)}"