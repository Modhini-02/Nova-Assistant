# import google.generativeai as genai

# genai.configure(api_key="AIzaSyBQn0U844Y3uLyKtp2sdA33yI7MfQLS2eU")
# models = genai.list_models()
# for m in models:
#     print(m)
# import os
# path = "C:\\Users\\yaswa\\Music"  # Replace with your music folder path
# print("Music library exists:", os.path.exists(path))
import subprocess
MEDIA_PLAYER = "wmplayer.exe"  # or "music.ui.exe" if that's what you see
output = subprocess.check_output('tasklist', shell=True).decode().lower().splitlines()
found = any(MEDIA_PLAYER.lower() in line for line in output)
print("Media player running:", found)
