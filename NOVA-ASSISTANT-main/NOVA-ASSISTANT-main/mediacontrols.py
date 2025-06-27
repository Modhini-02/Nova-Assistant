import pyautogui
import os
import random

# Set your music library path here:
MUSIC_LIBRARY_PATH = r"C:\Users\YourUsername\Music"  # <-- Replace with your actual music folder path.

# === Coordinates for media control buttons ===
# Update these coordinates based on your media player's UI location.
# Coordinates for the Play/Pause button:
PLAY_PAUSE_BUTTON = (1016, 999)      # Example: (x, y) for Play/Pause button
# Coordinates for the Next button:
NEXT_BUTTON = (1074, 999)            # Example: (x, y) for Next track button
# Coordinates for the Previous button:
PREVIOUS_BUTTON = (957, 999)         # Example: (x, y) for Previous track button
# Coordinates for the Close button (if you want to close the media player window):
CLOSE_BUTTON = (1745, 104)           # Example: (x, y) for the window’s "X" button

def control_media(action):
    if action == "play":
        # If media player is not open, open a random song.
        if not is_media_player_open():
            result = play_song_from_library()
            return result
        # Click on the Play/Pause button.
        pyautogui.click(PLAY_PAUSE_BUTTON[0], PLAY_PAUSE_BUTTON[1])
        return "Play/Pause button clicked."
    
    elif action == "pause":
        if is_media_player_open():
            pyautogui.click(PLAY_PAUSE_BUTTON[0], PLAY_PAUSE_BUTTON[1])
            return "Play/Pause button clicked (pause)."
        else:
            return "Media player is not open."
    
    elif action == "next":
        pyautogui.click(NEXT_BUTTON[0], NEXT_BUTTON[1])
        return "Next track button clicked."
    
    elif action == "previous":
        pyautogui.click(PREVIOUS_BUTTON[0], PREVIOUS_BUTTON[1])
        return "Previous track button clicked."
    
    elif action == "stop":
        # Instead of force-killing the process (which can cause hangs),
        # simulate a click on the close button of the media player's window.
        pyautogui.click(CLOSE_BUTTON[0], CLOSE_BUTTON[1])
        return "Close button clicked on media player."
    
    elif action == "volume up":
        # If you have coordinates for volume up, update the following:
        # Example: VOLUME_UP_BUTTON = (x, y)
        # pyautogui.click(VOLUME_UP_BUTTON[0], VOLUME_UP_BUTTON[1])
        return "Volume up command not implemented via coordinates."
    
    elif action == "volume down":
        # Similarly, if you have coordinates for volume down:
        return "Volume down command not implemented via coordinates."
    
    else:
        return "Unsupported action."

def is_media_player_open():
    """
    Checks if the media player window is open.
    For a coordinate-based approach, you can use image recognition (pyautogui.locateOnScreen)
    or set a flag when you open the media player.
    
    For simplicity, this function currently returns True if the music library path exists,
    but you may want to implement a more robust check.
    """
    # As a placeholder, assume the media player is open if a recent song was played.
    # You can enhance this function using pygetwindow or image recognition.
    return True

def play_song_from_library():
    """Plays a random song from your music library using os.startfile()."""
    try:
        if not os.path.exists(MUSIC_LIBRARY_PATH):
            return f"Music library path not found: {MUSIC_LIBRARY_PATH}"
        # List music files with common extensions.
        music_files = [f for f in os.listdir(MUSIC_LIBRARY_PATH) if f.lower().endswith(('.mp3', '.wav', '.flac'))]
        if not music_files:
            return "No music files found in the library."
        random_song = random.choice(music_files)
        song_path = os.path.join(MUSIC_LIBRARY_PATH, random_song)
        os.startfile(song_path)  # This opens the file with the default media player.
        return f"Playing {random_song}"
    except Exception as e:
        return f"Failed to play song: {e}"
