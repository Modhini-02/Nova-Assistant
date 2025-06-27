import time
import webbrowser
import pyautogui
import pygetwindow as gw
import speech_recognition as sr

# ------------------ Helper for Voice Input ------------------

def get_voice_input(prompt):
    """
    Prompts the user (via console) and listens for a voice response.
    Returns the recognized text or an empty string if recognition fails.
    """
    print(prompt)
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
    try:
        text = r.recognize_google(audio)
        print("You said:", text)
        return text.strip()
    except Exception as e:
        print("Voice input error:", e)
        return ""

# ------------------ Gmail Opening and Reading ------------------

def open_gmail():
    """
    Opens Gmail in the default web browser and waits for it to load.
    """
    webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
    time.sleep(10)  # Adjust sleep time as needed.
    return "Gmail opened."

def read_mail(n):
    """
    Clicks on the nth email in the Gmail inbox using fixed coordinates.
    Update base_x, base_y, and offset according to your screen layout.
    """
    try:
        n = int(n)
    except Exception:
        return "Invalid mail number provided."
    
    # Example coordinates for the first email; adjust these for your screen.
    base_x = 855
    base_y = 335
    offset = (n - 1) * 50  # Adjust vertical spacing if necessary

    pyautogui.click(base_x, base_y + offset)
    time.sleep(5)
    return f"Opened mail number {n}."

def close_gmail():
    """
    Closes the Gmail window. It now searches for any window with 'gmail' (case-insensitive) in its title.
    If no such window is found, it falls back to using the Ctrl+W hotkey.
    """
    try:
        gmail_windows = [window for window in gw.getAllWindows() if "gmail" in window.title.lower()]
        if gmail_windows:
            for window in gmail_windows:
                window.close()
            time.sleep(3)
            return "Gmail closed."
        else:
            pyautogui.hotkey("ctrl", "w")
            time.sleep(2)
            return "Gmail closed via hotkey."
    except Exception as e:
        return f"Error closing Gmail: {str(e)}"

# ------------------ Coordinate-Based Sending ------------------

def compose_mail():
    """
    Clicks the Compose button using fixed coordinates.
    Update compose_x and compose_y to the actual location of the Compose button.
    """
    compose_x = 42   # <-- Update these values as per your screen!
    compose_y = 225  # <-- Update these values as per your screen!
    pyautogui.click(compose_x, compose_y)
    time.sleep(2)
    return True

def click_send():
    """
    Clicks the Send button using fixed coordinates.
    Update send_x and send_y to the actual location of the Send button.
    """
    send_x = 1145   # <-- Update these values as per your screen!
    send_y = 1095   # <-- Update these values as per your screen!
    pyautogui.click(send_x, send_y)
    time.sleep(2)

def send_mail(recipient, subject, body):
    """
    Automates sending an email in Gmail using fixed coordinates.
    It clicks Compose, writes the recipient, subject, and body,
    and then clicks Send.
    """
    if not compose_mail():
        return "Compose button not found. Please verify the coordinates."
    
    # Write recipient; assume the "To" field is focused.
    pyautogui.write(recipient, interval=0.05)
    pyautogui.press("tab")  # Move to Subject field.
    time.sleep(0.5)
    
    pyautogui.write(subject, interval=0.05)
    # Press Tab twice to ensure focus moves to the body (adjust if necessary)
    pyautogui.press("tab", presses=2, interval=0.3)
    time.sleep(0.5)
    
    pyautogui.write(body, interval=0.05)
    time.sleep(0.5)
    
    click_send()
    return "Email sent successfully."

def handle_gmail_send_command(user_input):
    """
    Processes a 'compose mail' or 'send mail' command. If details aren't included,
    it will prompt for them via voice input. The recipient is taken as a name;
    if no '@' is found, '@gmail.com' is appended.
    
    Expected initial command example: "compose mail" or "send mail"
    """
    # Remove both "send mail" and "compose mail" prefixes.
    command = user_input.lower().replace("send mail", "").replace("compose mail", "").strip()
    recipient = None
    subject = None
    body = None

    # Use voice input for each field.
    if not recipient:
        recipient = get_voice_input("Please say the recipient's name (provide only the part before '@' for Gmail):")
    if not subject:
        subject = get_voice_input("Please say the subject of the email:")
    if not body:
        body = get_voice_input("Please say the body of the email:")

    # Append '@gmail.com' if missing.
    if "@" not in recipient:
        recipient = recipient + "@gmail.com"
    
    return send_mail(recipient, subject, body)

# ------------------ Coordinate-Based Searching ------------------

def search_mail(query):
    """
    Searches for emails in Gmail using the search box.
    Clicks on the search box at fixed coordinates, types the query, and presses Enter.
    Update search_box_x and search_box_y to match your Gmail layout.
    """
    search_box_x = 400  # <-- Update with your search box's X coordinate.
    search_box_y = 150  # <-- Update with your search box's Y coordinate.
    
    pyautogui.click(search_box_x, search_box_y)
    time.sleep(1)
    pyautogui.write(query, interval=0.05)
    pyautogui.press("enter")
    time.sleep(5)
    return f"Displayed search results for '{query}'."
def gmail_home():
    """
    Clicks the Gmail home icon using fixed coordinates.
    Update home_x and home_y to the actual location of the Gmail name icon at the top left.
    """
    home_x = 100   # <-- Update these values as per your screen!
    home_y = 50    # <-- Update these values as per your screen!
    pyautogui.click(home_x, home_y)
    time.sleep(3)
    return "Gmail home clicked."

