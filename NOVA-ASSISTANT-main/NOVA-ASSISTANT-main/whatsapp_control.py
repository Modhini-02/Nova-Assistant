# whatsapp_control.py
import os
import subprocess
import time
import pyautogui
import pygetwindow as gw
import winreg
import urllib.parse
import psutil
def get_whatsapp_path():
    """Check for WhatsApp Desktop via registry or known installation paths."""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WhatsAppDesktop.exe") as key:
            store_path = winreg.QueryValue(key, None)
            if os.path.exists(store_path):
                return store_path
    except Exception:
        pass

    # Fallback: Look in WindowsApps folder (Microsoft Store installations)
    base_path = r"C:\Program Files\WindowsApps"
    package_pattern = "5319275A.WhatsAppDesktop_"
    try:
        versions = [d for d in os.listdir(base_path) if d.startswith(package_pattern)]
        if versions:
            versions.sort(reverse=True)
            latest_path = os.path.join(base_path, versions[0], "WhatsAppDesktop.exe")
            if os.path.exists(latest_path):
                return latest_path
    except Exception:
        pass

    return None

def is_whatsapp_installed():
    """Return True if WhatsApp Desktop appears to be installed or is running."""
    path = get_whatsapp_path()
    if path is not None:
        return True
    wins = gw.getWindowsWithTitle("WhatsApp")
    return len(wins) > 0

def open_whatsapp():
    """
    Opens WhatsApp Desktop.
    - If WhatsApp Desktop is installed, it attempts to open it via its URI scheme.
    - Otherwise, it returns an error message.
    """
    if is_whatsapp_installed():
        ret = os.system("start whatsapp:")
        if ret == 0:
            time.sleep(10)  # wait for the app to load
            return "WhatsApp Desktop opened."
        else:
            # Fallback: use the direct executable if possible
            whatsapp_path = get_whatsapp_path()
            if whatsapp_path:
                subprocess.Popen([whatsapp_path])
                time.sleep(3)
                return "WhatsApp Desktop opened via direct launch."
            return "Failed to open WhatsApp Desktop via protocol."
    else:
        return "WhatsApp Desktop not installed."

def focus_whatsapp():
    """Bring the WhatsApp window to the front with a few retries."""
    for _ in range(3):
        wins = gw.getWindowsWithTitle("WhatsApp")
        if wins:
            win = wins[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            time.sleep(2)
            return True
        time.sleep(1)
    return False

def open_chat_with_contact(contact_name):
    """
    Opens the chat with the given contact name using WhatsApp Desktop automation.
    Uses Ctrl+F to open search, types the contact name, then selects the first result.
    """
    if not is_whatsapp_installed():
        return "WhatsApp Desktop not installed."
    if not focus_whatsapp():
        open_whatsapp()
        if not focus_whatsapp():
            return "Couldn't bring WhatsApp Desktop to focus."

    try:
        # Open search box
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        # Type the contact name
        pyautogui.write(contact_name, interval=0.1)
        time.sleep(1)
        # Select the first result
        pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(2)
        return f"Opened chat with {contact_name}"
    except Exception as e:
        return f"Error opening chat: {str(e)}"

def send_whatsapp_message(message):
    """
    Sends a message in the currently active WhatsApp chat.
    Uses a fixed region click based on image recognition to locate the input box.
    If you already have focus on the chat, you may simply click on a known location.
    """
    try:
        # If you want to improve reliability, you could capture an image of the input area.
        # For now, we assume the chat input is at a consistent location.
        pyautogui.click(1000, 970)
        time.sleep(0.5)
        pyautogui.write(message, interval=0.05)
        pyautogui.press('enter')
        time.sleep(1)
        return True
    except Exception:
        return False

def send_message_to_contact(contact_name, message, open_app=True):
    """
    Complete flow: if open_app is True, ensure WhatsApp is open,
    then open the chat for contact_name and send the message.
    """
    responses = []
    if open_app:
        resp = open_whatsapp()
        responses.append(resp)
        time.sleep(3)
    chat_resp = open_chat_with_contact(contact_name)
    responses.append(chat_resp)
    if "Error" not in chat_resp and "not installed" not in chat_resp:
        if send_whatsapp_message(message):
            responses.append(f"Sent message: {message}")
        else:
            responses.append("Failed to send message.")
    else:
        responses.append("Chat could not be opened.")
    return "\n".join(responses)

def send_message_to_number(phone, message):
    """
    Sends a message to a phone number using WhatsApp's URL scheme.
    """
    encoded_message = urllib.parse.quote(message)
    if is_whatsapp_installed():
        url = f"whatsapp://send?phone={phone}&text={encoded_message}"
        ret = os.system(f"start {url}")
        if ret == 0:
            return f"Message sent to {phone} via WhatsApp Desktop."
        else:
            return f"Failed to send message to {phone} via WhatsApp Desktop."
    else:
        return "WhatsApp Desktop not installed."

# ----- New Features using Image Recognition -----

def start_video_call():
    """
    Initiates a video call on the currently active WhatsApp chat using image recognition.
    Assumes that 'video_call.png' is a screenshot of the video call button.
    """
    try:
        video_call_image = 'video_call.png'
        # Locate the center of the video call button on the screen.
        button_location = pyautogui.locateCenterOnScreen(video_call_image, confidence=0.8)
        if button_location:
            pyautogui.click(button_location)
            time.sleep(2)
            return "Video call initiated."
        else:
            return "Video call button not found on screen. Please check the image."
    except Exception as e:
        return f"Error initiating video call: {str(e)}"

def start_voice_call():
    """
    Initiates a voice call on the currently active WhatsApp chat using image recognition.
    Assumes that 'voice_call.png' is a screenshot of the voice call button.
    """
    try:
        voice_call_image = 'voice_call.png'
        button_location = pyautogui.locateCenterOnScreen(voice_call_image, confidence=0.8)
        if button_location:
            pyautogui.click(button_location)
            time.sleep(2)
            return "Voice call initiated."
        else:
            return "Voice call button not found on screen. Please check the image."
    except Exception as e:
        return f"Error initiating voice call: {str(e)}"

def view_status():
    """
    Opens the Status tab in WhatsApp Desktop using image recognition.
    Assumes that 'status.png' is a screenshot of the status tab button.
    """
    try:
        status_image = 'status.png'
        button_location = pyautogui.locateCenterOnScreen(status_image, confidence=0.8)
        if button_location:
            pyautogui.click(button_location)
            time.sleep(2)
            return "Status tab opened."
        else:
            return "Status button not found on screen. Please check the image."
    except Exception as e:
        return f"Error opening status tab: {str(e)}"
def mute_current_chat():
    """
    Mutes the currently selected chat in the WhatsApp left sidebar by:
    1) Right-clicking the chat area in the left sidebar.
    2) Clicking the 'Mute notifications' option in the context menu.
    3) (Optional) Selecting a mute duration (e.g., 8 hours).
    """
    try:
        # 1. Locate and right-click on the currently open chat in the left sidebar.
        #    e.g., 'current_chat_left_side.png' is a small screenshot of the contact name or chat area in the sidebar.
        chat_area = pyautogui.locateCenterOnScreen('current_chat_left_side.png', confidence=0.8)
        if chat_area:
            pyautogui.rightClick(chat_area)
            time.sleep(1)
        else:
            return "Could not find the chat in the left sidebar to right-click."

        # 2. Locate 'Mute notifications' in the context menu (e.g., 'mute_notifications.png').
        mute_option = pyautogui.locateCenterOnScreen('mute_notifications.png', confidence=0.8)
        if mute_option:
            pyautogui.click(mute_option)
            time.sleep(1)
        else:
            return "Could not find 'Mute notifications' in the context menu."

        # 3. (Optional) If a submenu with durations appears, locate and click the desired duration.
        #    For example, 'mute_8_hours.png'.
        # duration_option = pyautogui.locateCenterOnScreen('mute_8_hours.png', confidence=0.8)
        # if duration_option:
        #     pyautogui.click(duration_option)
        #     time.sleep(1)

        return "Chat muted successfully."
    except Exception as e:
        return f"Error muting chat: {str(e)}"


def unmute_current_chat():
    """
    Unmutes the currently selected chat in the WhatsApp left sidebar by:
    1) Right-clicking the chat area in the left sidebar.
    2) Clicking the 'Unmute notifications' option in the context menu.
    """
    try:
        # 1. Locate and right-click on the currently open chat in the left sidebar.
        chat_area = pyautogui.locateCenterOnScreen('current_chat_left_side.png', confidence=0.8)
        if chat_area:
            pyautogui.rightClick(chat_area)
            time.sleep(1)
        else:
            return "Could not find the chat in the left sidebar to right-click."

        # 2. Locate 'Unmute notifications' in the context menu (e.g., 'unmute_notifications.png').
        unmute_option = pyautogui.locateCenterOnScreen('unmute_notifications.png', confidence=0.8)
        if unmute_option:
            pyautogui.click(unmute_option)
            time.sleep(1)
        else:
            return "Could not find 'Unmute notifications' in the context menu."

        return "Chat unmuted successfully."
    except Exception as e:
        return f"Error unmuting chat: {str(e)}"
def close_whatsapp():
    """Closes WhatsApp Desktop if it is running."""
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if "WhatsApp.exe" in process.info['name']:
            try:
                psutil.Process(process.info['pid']).terminate()
                return "WhatsApp Desktop closed."
            except Exception as e:
                return f"Error closing WhatsApp: {str(e)}"
    return "WhatsApp is not running."