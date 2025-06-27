import os
import subprocess
import cv2
import numpy as np
import mss
import pyautogui
import time

def click_turn_on_now_opencv(template_path="turn_on_now.png", threshold=0.8, timeout=10):
    """
    Wait up to `timeout` seconds for the template image to appear on the screen.
    Once found, click the center of the detected area.
    
    Args:
        template_path (str): Path to the template image (screenshot of the "Turn on now" button).
        threshold (float): Matching threshold (0.0 to 1.0). Higher is more strict.
        timeout (int): Maximum seconds to wait before timing out.
        
    Returns:
        str: Result message.
    """
    # Load the template image using OpenCV
    template = cv2.imread(template_path)
    if template is None:
        return f"Template image '{template_path}' not found. Check your file path."
    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
    w, h = template_gray.shape[::-1]
    
    start_time = time.time()
    
    with mss.mss() as sct:
        while True:
            # Capture full screen (monitor 1)
            screenshot = sct.grab(sct.monitors[1])
            # Convert the raw data to a NumPy array
            img = np.array(screenshot)
            # Convert from BGRA to grayscale for matching
            img_gray = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
            
            # Perform template matching
            res = cv2.matchTemplate(img_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            loc = np.where(res >= threshold)
            
            if len(loc[0]) > 0:
                # Get the first matching location
                pt = (loc[1][0], loc[0][0])
                center_x = pt[0] + w // 2
                center_y = pt[1] + h // 2
                # Click at the center of the found region
                pyautogui.click(center_x, center_y)
                return "Clicked 'Turn on now' button using OpenCV."
            
            # Check timeout
            if time.time() - start_time > timeout:
                return "Timeout: Could not locate 'Turn on now' button."
            
            time.sleep(0.5)

def open_night_light_settings():
    """
    Open the Night Light settings page (no auto-toggle).
    """
    os.system("start ms-settings:nightlight")
    return "Opened Night Light settings."

def enable_night_light():
    """
    Attempt to enable Night Light via registry.
    May require logoff/logon to fully take effect.
    (This path may not exist on all Windows builds.)
    """
    try:
        commands = [
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Cloud\SystemSettings_Display_BlueLightReductionState -Name Data -Type Binary -Value 02000000"',
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Cloud\SystemSettings_Display_BlueLightReductionSettings -Name Data -Type Binary -Value 01000000000000000000000000003C42"'
        ]
        for cmd in commands:
            subprocess.run(cmd, shell=True, check=True)
        return "Night Light enabled (may require logoff/logon)."
    except Exception as e:
        return f"Error enabling Night Light: {e}"

def disable_night_light():
    """
    Attempt to disable Night Light via registry.
    May require logoff/logon to fully take effect.
    (This path may not exist on all Windows builds.)
    """
    try:
        commands = [
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Cloud\SystemSettings_Display_BlueLightReductionState -Name Data -Type Binary -Value 01000000"',
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Cloud\SystemSettings_Display_BlueLightReductionSettings -Name Data -Type Binary -Value 01000000000000000000000000003C00"'
        ]
        for cmd in commands:
            subprocess.run(cmd, shell=True, check=True)
        return "Night Light disabled (may require logoff/logon)."
    except Exception as e:
        return f"Error disabling Night Light: {e}"

# --------------------------------------------------------------------------
# DARK MODE
# --------------------------------------------------------------------------
def enable_dark_mode():
    """
    Set both system and app theme to Dark Mode.
    """
    try:
        commands = [
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Type DWord -Value 0"',
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Type DWord -Value 0"'
        ]
        for cmd in commands:
            subprocess.run(cmd, shell=True, check=True)
        return "Dark Mode enabled."
    except Exception as e:
        return f"Error enabling Dark Mode: {e}"

def disable_dark_mode():
    """
    Set both system and app theme to Light Mode.
    """
    try:
        commands = [
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Type DWord -Value 1"',
            r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Type DWord -Value 1"'
        ]
        for cmd in commands:
            subprocess.run(cmd, shell=True, check=True)
        return "Light Mode enabled."
    except Exception as e:
        return f"Error disabling Dark Mode: {e}"

# --------------------------------------------------------------------------
# FOCUS ASSIST
# --------------------------------------------------------------------------
def set_focus_assist_off():
    """
    Turn Focus Assist OFF (all notifications).
    Registry value: 0
    """
    try:
        cmd = r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings -Name NOC_GLOBAL_SETTING_TOASTS_ENABLED -Type DWord -Value 0"'
        subprocess.run(cmd, shell=True, check=True)
        return "Focus Assist turned off."
    except Exception as e:
        return f"Error turning off Focus Assist: {e}"

def set_focus_assist_priority():
    """
    Focus Assist to PRIORITY ONLY.
    Registry value: 1
    """
    try:
        cmd = r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings -Name NOC_GLOBAL_SETTING_TOASTS_ENABLED -Type DWord -Value 1"'
        subprocess.run(cmd, shell=True, check=True)
        return "Focus Assist set to Priority Only."
    except Exception as e:
        return f"Error setting Focus Assist to Priority Only: {e}"

def set_focus_assist_alarms():
    """
    Focus Assist to ALARMS ONLY.
    Registry value: 2
    """
    try:
        cmd = r'powershell -Command "Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings -Name NOC_GLOBAL_SETTING_TOASTS_ENABLED -Type DWord -Value 2"'
        subprocess.run(cmd, shell=True, check=True)
        return "Focus Assist set to Alarms Only."
    except Exception as e:
        return f"Error setting Focus Assist to Alarms Only: {e}"
