
from file_operations import create_code_file
from mediacontrols import control_media
from play_video_on_youtube import play_youtube_video
from run_file import run_specific_file
from scroll import scrolldown, scrollup
from voice_handler import listen, speak
from context_manager_error import generate_contextual_response
from database_handler import reset_conversation_history
from app_control import open_application

from database_handler import reset_conversation_history
from app_control import close_application, open_application
from browser_control import search_google
from write_to_file import write_code_in_vscode
from youtube_control import control_youtube

def main(log_function):
    user_id = "user_001"
    # reset_conversation_history(user_id)

    speak("Hello! I'm Nova. How can I assist you today?")
    log_function("Nova: Hello! I'm Nova. How can I assist you today?")

    while True:
        user_input = listen()
        log_function(f"You: {user_input}")

        if user_input.lower().startswith("timeout") or "didn't catch" in user_input.lower():
            speak(user_input)
            continue

        if any(cmd in user_input.lower() for cmd in ["exit", "quit"]):
            speak("Goodbye!")
            log_function("Nova: Goodbye!")
            break

        # Handle application and website commands
        if "open" in user_input.lower():
            response = open_application(user_input)
            speak(response)
            log_function(f"Nova: {response}")
            continue

        if "close" in user_input.lower():
            response = close_application(user_input)
            log_function(f"Nova: {response}")
            continue

        # Handle search commands
        if "search" in user_input.lower():
            response = search_google(user_input)
            speak(response)
            log_function(f"Nova: {response}")
            continue

        if "create" in user_input.lower():
            response=create_code_file(user_input)
            log_function(response)
            continue

        if "write" in user_input.lower() and "in" in user_input.lower():
            words = user_input.replace("write", "").strip().split(" in ")
            if len(words) == 2:
                code_description = words[0].strip()  # Extract code description
                language = words[1].strip()  # Extract programming language
                write_code_in_vscode(code_description, language)  #Calls correct function
            else:
                print("Please specify the code and language. Example: 'Write bubble sort in Python'.")
            continue
        
        if "run" in user_input.lower():
            words = user_input.replace("run", "").strip().split()
            if len(words) >= 3 and words[-1] == "file":  # Ensure last word is "file"
                filename = words[-3]  # Third last word is filename
                file_type = words[-2]  # Second last word is the actual type

                response=run_specific_file(filename, file_type)  # Call function with correct values
                log_function(f" Nova: {response}")
            else:
                print(" Please specify the filename and file type correctly. Example: 'Run new Python file'.")
            continue
        
        if "scroll up" in user_input.lower():
            scrollup()
            continue
        if "scroll down" in user_input.lower():
            scrolldown()
            continue

        if "play" in user_input.lower() and "on youtube" in user_input.lower():
            video_name = user_input.replace("play", "").replace("on youtube", "").strip()
            play_youtube_video(video_name)
            continue

        if "pause youtube" in user_input.lower() or "play youtube" in user_input.lower():
            response = control_youtube("pause" if "pause" in user_input.lower() else "play")
            log_function(f"Nova: {response}")
            continue
        
        try:
            if "pause music" in user_input.lower() or "play music" in user_input.lower():
                response = control_media("pause")  # ✅ Toggle Play/Pause
                log_function(f"Nova: {response}")
                continue

            if "next song" in user_input.lower():
                response = control_media("next")  # ✅ Skip to next track
                log_function(f"Nova: {response}")
                continue

            if "previous song" in user_input.lower():
                response = control_media("previous")  # ✅ Play previous track
                log_function(f"Nova: {response}")
                continue

            if "stop music" in user_input.lower():
                response = control_media("stop")  # ✅ Stop playback
                log_function(f"Nova: {response}")
                continue

            if "increase volume" in user_input.lower():
                response = control_media("volume up")  # 🔊 Increase volume
                log_function(f"Nova: {response}")
                continue

            if "decrease volume" in user_input.lower():
                response = control_media("volume down")  # 🔉 Decrease volume
                log_function(f"Nova: {response}")
                continue

        except Exception as e:
            log_function(f"Nova: Error: {e}")  # ✅ Log the error without stopping
            print(f" Error: {e}")
            continue
        # Handle other commands
        nova_response = generate_contextual_response(user_input, user_id)
        speak(nova_response)
        log_function(f"Nova: {nova_response}")
