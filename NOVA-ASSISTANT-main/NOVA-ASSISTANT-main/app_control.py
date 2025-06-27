import os
import pyautogui
import time
import subprocess
import pygetwindow as gw
import webbrowser
from whatsapp_control import open_whatsapp
from gmail_control import open_gmail
from mediaplayer import close_media_app, open_media_app

# Define sample coordinates for the close button of various applications.
# UPDATE these values to match your screen.
CLOSE_BUTTON_COORDINATES = {
    "notepad": (1894, 15),       # e.g., top-right 'X' of Notepad window
    "calculator": (1894, 15),    # update as needed
    "file explorer": (1894, 15), # update as needed
    "vs code": (1894, 15),       # update as needed
}

def open_application(command):
    if "open gmail" in command:
        return open_gmail()
    if "open whatsapp" in command:
        return open_whatsapp()
    if "notepad" in command:
        subprocess.Popen("notepad.exe")
        return "Notepad opened."
    elif "calculator" in command:
        os.system("calc")
        return "Calculator opened."
    elif "file explorer" in command or "open explorer" in command:
        os.system("start explorer")
        return "File Explorer opened."
    elif "open vs code" in command or "open visual studio code" in command:
        VS_CODE_FOLDER = "D:/Nova-testing"
        os.system(f"code {VS_CODE_FOLDER}")
        return f"Opened VS Code in {VS_CODE_FOLDER}"
    elif "open google" in command or "open browser" in command:
        try:
            webbrowser.get("chrome").open("https://www.google.com")
            return "Opened Google in Chrome."
        except webbrowser.Error:
            os.system("start msedge https://www.google.com")
            return "Opened Google in Edge."
    elif "open youtube" in command:
        webbrowser.open("https://www.youtube.com")
        return "Opened YouTube."
    elif "open twitter" in command:
        webbrowser.open("https://www.twitter.com")
        return "Opened Twitter."
    elif "open instagram" in command:
        webbrowser.open("https://www.instagram.com")
        return "Opened Instagram."
    elif "open facebook" in command:
        webbrowser.open("https://www.facebook.com")
        return "Opened Facebook."
    elif "open pictures" in command:
        os.startfile(os.path.expanduser("~/Pictures"))
        return "Opened Pictures folder."
    elif "open documents" in command:
        os.startfile(os.path.expanduser("~/Documents"))
        return "Opened Documents folder."
    elif "open downloads" in command:
        os.startfile(os.path.expanduser("~/Downloads"))
        return "Opened Downloads folder."
    elif "open music folder" in command:
        os.startfile(os.path.expanduser("~/Music"))
        return "Opened Music folder."
    elif "open music" in command or "open video" in command:
        open_media_app()
        return "Opened Media Player."
    else:
        return "Application or command not recognized."

def close_application(command):
    command_lower = command.lower()
    # Use coordinate-based clicks if coordinates are defined.
    if "notepad" in command_lower:
        coords = CLOSE_BUTTON_COORDINATES.get("notepad")
        if coords:
            pyautogui.click(coords[0], coords[1])
            return "Closed Notepad via coordinate click."
        else:
            pyautogui.hotkey("alt", "f4")
            return "Closed Notepad via Alt+F4."
    elif "calculator" in command_lower:
        coords = CLOSE_BUTTON_COORDINATES.get("calculator")
        if coords:
            pyautogui.click(coords[0], coords[1])
            return "Closed Calculator via coordinate click."
        else:
            pyautogui.hotkey("alt", "f4")
            return "Closed Calculator via Alt+F4."
    elif "file explorer" in command_lower or "explorer" in command_lower:
        coords = CLOSE_BUTTON_COORDINATES.get("file explorer")
        if coords:
            pyautogui.click(coords[0], coords[1])
            return "Closed File Explorer via coordinate click."
        else:
            pyautogui.hotkey("alt", "f4")
            return "Closed File Explorer via Alt+F4."
    elif "vs code" in command_lower or "visual studio code" in command_lower:
        coords = CLOSE_BUTTON_COORDINATES.get("vs code")
        if coords:
            pyautogui.click(coords[0], coords[1])
            return "Closed VS Code via coordinate click."
        else:
            pyautogui.hotkey("alt", "f4")
            return "Closed VS Code via Alt+F4."
    elif "browser" in command_lower or "visual studio code" in command_lower:
        coords = CLOSE_BUTTON_COORDINATES.get("vs code")
        if coords:
            pyautogui.click(coords[0], coords[1])
            return "Closed browser via coordinate click."
        else:
            
            # For browsers, close the current tab using Ctrl+W.
            pyautogui.hotkey("ctrl", "w")
            return "Closed browser tab via Ctrl+W."
    else:
        return "Application not recognized for closing."
