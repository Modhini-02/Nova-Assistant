import os

def open_display_settings():
    """
    Open Windows Display Settings.
    """
    os.system("start ms-settings:display")
    return "Opened Display Settings."

def open_night_light_settings():
    """
    Open Windows Night Light Settings.
    """
    os.system("start ms-settings:nightlight")
    return "Opened Night Light Settings."
