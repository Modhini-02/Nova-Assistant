import os
import platform
import subprocess
import screen_brightness_control as sbc

# Function to adjust brightness
def set_brightness(level):
    try:
        if platform.system() == "Windows":
            sbc.set_brightness(level)
            return f"Brightness set to {level}%."
        else:
            return "Brightness control is only supported on Windows."
    except Exception as e:
        return f"Error setting brightness: {str(e)}"

# Function to increase brightness
def increase_brightness():
    try:
        if platform.system() == "Windows":
            current_brightness = sbc.get_brightness()[0]
            new_brightness = min(current_brightness + 10, 100)
            sbc.set_brightness(new_brightness)
            return f"Brightness increased to {new_brightness}%."
        else:
            return "Brightness control is only supported on Windows."
    except Exception as e:
        return f"Error increasing brightness: {str(e)}"

# Function to decrease brightness
def decrease_brightness():
    try:
        if platform.system() == "Windows":
            current_brightness = sbc.get_brightness()[0]
            new_brightness = max(current_brightness - 10, 0)
            sbc.set_brightness(new_brightness)
            return f"Brightness decreased to {new_brightness}%."
        else:
            return "Brightness control is only supported on Windows."
    except Exception as e:
        return f"Error decreasing brightness: {str(e)}"

# Function to restart the system
def restart_system():
    try:
        if platform.system() == "Windows":
            os.system("shutdown /r /t 1")
            return "Restarting the system..."
        else:
            return "Restart is only supported on Windows."
    except Exception as e:
        return f"Error restarting system: {str(e)}"

# Function to shut down the system
def shutdown_system():
    try:
        if platform.system() == "Windows":
            os.system("shutdown /s /t 1")
            return "Shutting down the system..."
        else:
            return "Shutdown is only supported on Windows."
    except Exception as e:
        return f"Error shutting down system: {str(e)}"

# Function to lock the system
def lock_system():
    try:
        if platform.system() == "Windows":
            os.system("rundll32.exe user32.dll,LockWorkStation")
            return "System locked."
        else:
            return "Locking is only supported on Windows."
    except Exception as e:
        return f"Error locking system: {str(e)}"