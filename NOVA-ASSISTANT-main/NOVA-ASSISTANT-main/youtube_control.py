# import pyautogui
# import time

# def search_youtube(query):
#     # Coordinates for YouTube's search box.
#     SEARCH_BOX_COORDINATES = (711, 144)  # Update these coordinates.
#     pyautogui.click(SEARCH_BOX_COORDINATES[0], SEARCH_BOX_COORDINATES[1])
#     time.sleep(1)
#     pyautogui.write(query, interval=0.05)
#     pyautogui.press("enter")
#     return f"Searched YouTube for: {query}"

# def click_video(n):
#     # Coordinates for the first video thumbnail.
#     BASE_VIDEO_COORDINATES = (690, 411)  # Update these coordinates.
#     VERTICAL_OFFSET = 400  # Adjust this offset based on your screen.
#     pyautogui.click(BASE_VIDEO_COORDINATES[0], BASE_VIDEO_COORDINATES[1] + (n - 1) * VERTICAL_OFFSET)
#     return f"Clicked video number {n}"

# def skip_ads():
#     # Coordinates for the "Skip Ad" button on YouTube.
#     SKIP_BUTTON_COORDINATES = (1261, 789)  # Update these coordinates.
#     pyautogui.click(SKIP_BUTTON_COORDINATES[0], SKIP_BUTTON_COORDINATES[1])
#     return "Skipped ad."

# def control_youtube(action):
#     if action == "play" or action == "pause":
#         # Coordinates for the YouTube play/pause button.
#         PLAY_BUTTON_COORDINATES = (699, 523)  # Update these coordinates.
#         pyautogui.click(PLAY_BUTTON_COORDINATES[0], PLAY_BUTTON_COORDINATES[1])
#         return f"YouTube {action} button clicked."
#     elif action == "skip ads":
#         return skip_ads()
#     elif action.startswith("click video"):
#         # Expected command format: "click video 3"
#         parts = action.split()
#         if len(parts) >= 3:
#             try:
#                 n = int(parts[2])
#                 return click_video(n)
#             except ValueError:
#                 return "Invalid video number."
#     else:
#         return f"Unsupported YouTube action: {action}"
import pyautogui
import time

def search_youtube(query):
    # Coordinates for YouTube's search box.
    SEARCH_BOX_COORDINATES = (729, 143)  # Update as needed.
    pyautogui.click(*SEARCH_BOX_COORDINATES)
    time.sleep(1)
    pyautogui.write(query, interval=0.05)
    pyautogui.press("enter")
    return f"Searched YouTube for: {query}"

def click_video(n):
    # Coordinates for the first video thumbnail.
    BASE_VIDEO_COORDINATES = (710, 411)  # Update as needed.
    VERTICAL_OFFSET = 400  # Adjust based on your screen.
    pyautogui.click(BASE_VIDEO_COORDINATES[0], BASE_VIDEO_COORDINATES[1] + (n - 1) * VERTICAL_OFFSET)
    return f"Clicked video number {n}"

def skip_ads():
    # Coordinates for the "Skip Ad" button.
    SKIP_BUTTON_COORDINATES = (1261, 789)  # Update as needed.
    pyautogui.click(*SKIP_BUTTON_COORDINATES)
    return "Skipped ad."

def play_pause():
    # Coordinates for the play/pause button.
    PLAY_BUTTON_COORDINATES = (699, 523)  # Update as needed.
    pyautogui.click(*PLAY_BUTTON_COORDINATES)
    return "YouTube play/pause button clicked."

def like_video():
    # Coordinates for the Like button.
    LIKE_BUTTON_COORDINATES = (752, 1022)  # Update as needed.
    pyautogui.click(*LIKE_BUTTON_COORDINATES)
    return "YouTube Like button clicked."

def dislike_video():
    # Coordinates for the Dislike button.
    DISLIKE_BUTTON_COORDINATES = (848, 1027)  # Update as needed.
    pyautogui.click(*DISLIKE_BUTTON_COORDINATES)
    return "YouTube Dislike button clicked."

def fullscreen():
    # Coordinates for the Full Screen button.
    FULLSCREEN_BUTTON_COORDINATES = (1288, 911)  # Update as needed.
    pyautogui.click(*FULLSCREEN_BUTTON_COORDINATES)
    return "YouTube Full Screen button clicked."

def next_video():
    # Coordinates for the Next Video button (if available in UI)
    NEXT_BUTTON_COORDINATES = (131, 911)  # Update as needed.
    pyautogui.click(*NEXT_BUTTON_COORDINATES)
    return "Next video button clicked."

def control_youtube(action):
    action = action.lower()
    if action in ["play", "pause"]:
        return play_pause()
    elif action == "skip ads":
        return skip_ads()
    elif action.startswith("click video"):
        # Expected format: "click video 3"
        parts = action.split()
        if len(parts) >= 3:
            try:
                n = int(parts[2])
                return click_video(n)
            except ValueError:
                return "Invalid video number."
    elif action == "like video":
        return like_video()
    elif action == "dislike video":
        return dislike_video()
    elif action == "full screen":
        return fullscreen()
    elif action == "next video":
        return next_video()
    elif action.startswith("search youtube"):
        query = action.replace("search youtube", "").strip()
        return search_youtube(query)
    else:
        return f"Unsupported YouTube action: {action}"
