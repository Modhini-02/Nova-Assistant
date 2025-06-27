from file_operations import create_code_file
from mediacontrols import control_media
from play_video_on_youtube import play_youtube_video
from run_file import run_specific_file
from scroll import scrolldown, scrollup
from voice_handler import listen, speak
from context_manager_error import generate_contextual_response
from database_handler import reset_conversation_history
from app_control import open_application, close_application
from browser_control import search_google,click_nth_link
from write_to_file import write_code_in_vscode
from youtube_control import control_youtube
import os
import re

class NovaAssistant:
    def __init__(self, user_id="user_001"):
        self.user_id = user_id
        # reset_conversation_history(self.user_id)

    def process_command(self, user_input, log_function, speak_response=True):
        lower_input = user_input.lower()

        if lower_input.startswith("timeout") or "didn't catch" in lower_input:
            response = user_input
            if speak_response:
                speak(response)
            log_function(f"Nova: {response}")
            return response

        if any(cmd in lower_input for cmd in ["exit", "quit"]):
            response = "Goodbye!"
            if speak_response:
                speak(response)
            log_function(f"Nova: {response}")
            return response

        # Handle application and website commands
        if "open" in lower_input:
            response = open_application(user_input)
            if speak_response:
                speak(response)
            log_function(f"Nova: {response}")
            return response
        if "close" in lower_input:
            response = close_application(user_input)
            log_function(f"Nova: {response}")
            return response

        # --- Browser Commands ---
        if "search google" in lower_input:
            response = search_google(lower_input)  # from browser_control.py
            log_function(f"Nova: {response}")
            return response

        elif "click link" in lower_input or "click nth link" in lower_input:
            import re
            match = re.search(r'\d+', lower_input)
            if match:
                n = int(match.group())
                response = click_nth_link(n)  # from browser_control.py
                log_function(f"Nova: {response}")
                return response
            else:
                log_function("Nova: No link number specified.")
                return "Please specify the link number to click."


        if "new file" in lower_input:
            response = create_code_file(user_input)
            log_function(response)
            return response

        if "write" in lower_input and "in" in lower_input:
            words = user_input.replace("write", "").strip().split(" in ")
            if len(words) == 2:
                code_description = words[0].strip()
                language = words[1].strip()
                write_code_in_vscode(code_description, language)
            else:
                log_function("Nova: Please specify the code and language. Example: 'Write bubble sort in Python'.")
            return None

        if "run" in lower_input:
            words = user_input.replace("run", "").strip().split()
            if len(words) >= 3 and words[-1] == "file":
                filename = words[-3]
                file_type = words[-2]
                response = run_specific_file(filename, file_type)
                log_function(f"Nova: {response}")
                return response
            else:
                log_function("Nova: Please specify the filename and file type correctly. Example: 'Run new Python file'.")
                return None

        if "scroll up" in lower_input:
            scrollup()
            return None
        if "scroll down" in lower_input:
            scrolldown()
            return None

        if "play" in lower_input and "on youtube" in lower_input:
            video_name = user_input.replace("play", "").replace("on youtube", "").strip()
            play_youtube_video(video_name)
            return None

        if "pause youtube" in lower_input or "play youtube" in lower_input:
            response = control_youtube("pause" if "pause" in lower_input else "play")
            log_function(f"Nova: {response}")
            return response
        # --- YouTube Commands ---
        if "search youtube" in lower_input:
            # Extract query after "search youtube"
           # query = lower_input.replace("search youtube", "").strip()
            # If query is empty, you might prompt the user via voice input
            if not lower_input:
                lower_input = "default query"  # or call get_voice_input() if desired
            response = control_youtube(lower_input)  # Alternatively, call youtube_control.search_youtube(query)
            log_function(f"Nova: {response}")
            return response

        elif "like youtube" in lower_input:
            response = control_youtube("like video")
            log_function(f"Nova: {response}")
            return response

        elif "dislike youtube" in lower_input:
            response = control_youtube("dislike video")
            log_function(f"Nova: {response}")
            return response
        elif "skip add" in lower_input:
            response = control_youtube("skip ads")
            log_function(f"Nova: {response}")
            return response

        elif "full screen youtube" in lower_input:
            response = control_youtube("full screen")
            log_function(f"Nova: {response}")
            return response

        elif "next video youtube" in lower_input:
            response = control_youtube("next video")
            log_function(f"Nova: {response}")
            return response

        elif "click video" in lower_input:
            # Expects command like "click video 3"
            response = control_youtube(lower_input)  # your youtube_control.control_youtube() parses the number
            log_function(f"Nova: {response}")
            return response

        try:
            if "play music" in lower_input:
                response = control_media("play")
                log_function(f"Nova: {response}")
                return response
            if "pause music" in lower_input:
                response = control_media("pause")
                log_function(f"Nova: {response}")
                return response
            if "next song" in lower_input:
                response = control_media("next")
                log_function(f"Nova: {response}")
                return response
            if "previous song" in lower_input:
                response = control_media("previous")
                log_function(f"Nova: {response}")
                return response
            if "stop music" in lower_input:
                response = control_media("stop")
                log_function(f"Nova: {response}")
                return response
            if "increase volume" in lower_input:
                response = control_media("volume up")
                log_function(f"Nova: {response}")
                return response
            if "decrease volume" in lower_input:
                response = control_media("volume down")
                log_function(f"Nova: {response}")
                return response
        except Exception as e:
            log_function(f"Nova: Error: {e}")
            print(f"Error: {e}")
            return None


        # Fallback: generate a contextual response
        nova_response = generate_contextual_response(user_input, self.user_id)
        if speak_response:
            speak(nova_response)
        log_function(f"Nova: {nova_response}")
        return nova_response

if __name__ == "__main__":
    def log_function(msg):
        print(msg)
    assistant = NovaAssistant()
    speak("Hello! I'm Nova. How can I assist you today?")
    log_function("Nova: Hello! I'm Nova. How can I assist you today?")
    while True:
        user_input = listen()
        log_function(f"You: {user_input}")
        if any(cmd in user_input.lower() for cmd in ["exit", "quit", "stop"]):
            speak("Goodbye!")
            log_function("Nova: Goodbye!")
            break
        assistant.process_command(user_input, log_function)
