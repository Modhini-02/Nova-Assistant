import ctypes
import subprocess

# Constants for WM_APPCOMMAND method (for relative adjustments)
WM_APPCOMMAND = 0x319
APPCOMMAND_VOLUME_UP   = 0x0a0000
APPCOMMAND_VOLUME_DOWN = 0x090000
APPCOMMAND_VOLUME_MUTE = 0x080000
NIRCMD_PATH = r"D:\mini-project\NOVA-ASSISTANT\nircmd-x64\nircmd.exe"
def set_volume(percent):
    if not (0 <= percent <= 100):
        return "Volume percentage must be between 0 and 100."
    
    vol = int(65535 * (percent / 100.0))
    try:
        subprocess.run([NIRCMD_PATH, "setsysvolume", str(vol)], check=True)
        return f"Volume set to {percent}%."
    except Exception as e:
        return f"Error setting volume: {e}"
def get_foreground_window():
    return ctypes.windll.user32.GetForegroundWindow()

def volume_up():
    hwnd = get_foreground_window()
    ctypes.windll.user32.SendMessageW(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_VOLUME_UP)
    return "Volume increased."

def volume_down():
    hwnd = get_foreground_window()
    ctypes.windll.user32.SendMessageW(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_VOLUME_DOWN)
    return "Volume decreased."

def toggle_mute():
    hwnd = get_foreground_window()
    ctypes.windll.user32.SendMessageW(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_VOLUME_MUTE)
    return "Toggled mute."
