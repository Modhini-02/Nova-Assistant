import webbrowser
import pyautogui
import time

#  Function to search and play a video on YouTube
def play_youtube_video(video_name):
    search_url = f"https://www.youtube.com/results?search_query={video_name.replace(' ', '+')}"
    webbrowser.open(search_url)  #  Open YouTube with the search query
    time.sleep(5)  # Wait for YouTube to load

    #  Click on the first video result (position may vary)
    pyautogui.moveTo(584, 802)  # Adjust the coordinates if needed
    pyautogui.click()

    print(f"Playing: {video_name} on YouTube")











































# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# import time

# #  Function to play a specific YouTube video
# def play_youtube_video(video_name):
#     driver = webdriver.Chrome()  # Make sure chromedriver.exe is installed
#     driver.get("https://www.youtube.com/")
#     time.sleep(3)  # Wait for YouTube to load

#     #  Find the search bar and enter the video name
#     search_box = driver.find_element(By.NAME, "search_query")
#     search_box.send_keys(video_name)
#     search_box.send_keys(Keys.RETURN)
#     time.sleep(3)  # Wait for results to load

#     #  Find the exact video based on title
#     videos = driver.find_elements(By.ID, "video-title")
#     for video in videos:
#         if video_name.lower() in video.text.lower():
#             print(f" Playing: {video.text}")
#             video.click()  # Click on the correct video
#             return

#     print(f"❌ Video '{video_name}' not found.")
#     driver.quit()

