import os
import subprocess

#  Common installation paths for media players
MEDIA_APP_PATHS = {
    "wmplayer": "C:\\Program Files\\Windows Media Player\\wmplayer.exe",
    "vlc": "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe",
    "groove": "C:\\Program Files\\WindowsApps\\Microsoft.ZuneMusic_8wekyb3d8bbwe\\Music.UI.exe",
    "spotify": "C:\\Users\\%USERNAME%\\AppData\\Roaming\\Spotify\\Spotify.exe",
    "itunes": "C:\\Program Files\\iTunes\\iTunes.exe",
}

#  Function to find the first available media player
def find_available_media_player():
    for app, path in MEDIA_APP_PATHS.items():
        if os.path.exists(path):
            return path  #  Return the first available media player
    return None  # No media player found

#  Function to open a media application
def open_media_app():
    app_path = find_available_media_player()

    if app_path:
        try:
            subprocess.Popen(app_path, shell=True)
            print(f" Opened {app_path}")
            return f"Opened {app_path}"
        except Exception as e:
            print(f" Error opening {app_path}: {e}")
            return f"Error opening {app_path}"
    else:
        print(" No media player found on this system.")
        return "No media player found."

#  Function to close the media application
def close_media_app():
    app_path = find_available_media_player()

    if app_path:
        try:
            if os.name == "nt":  #  Windows
                app_name = os.path.basename(app_path)  # Extracts filename from path
                subprocess.call(f"taskkill /IM {app_name} /F", shell=True)
            else:  #  Linux/Mac
                subprocess.call(f"pkill {app_path}", shell=True)

            print(f" Closed {app_path}")
            return f"Closed {app_path}"
        except Exception as e:
            print(f" Error closing {app_path}: {e}")
            return f"Error closing {app_path}"
    else:
        print(" No running media player found.")
        return "No running media player found."
