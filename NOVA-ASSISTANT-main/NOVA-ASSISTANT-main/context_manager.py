import os
import pyautogui
import keyboard
import time
from database_handler import get_conversation_history, update_conversation_history
from gemini_handler import generate_response
from file_manager import (
    open_directory, create_directory, rename_file, delete_file, list_files,
    navigate_to_path, open_file, open_application, open_file_with, copy_file, paste_file, move_file
)
from system_control import (
    set_brightness, increase_brightness, decrease_brightness,
    restart_system, shutdown_system, lock_system
)
from gmail_control import handle_gmail_send_command,open_gmail, read_mail, close_gmail,search_mail,get_voice_input,gmail_home
from whatsapp_control import (
 open_whatsapp, open_chat_with_contact, send_message_to_contact,
 send_message_to_number,send_whatsapp_message,start_video_call,view_status,start_voice_call,
 mute_current_chat,unmute_current_chat,close_whatsapp
) 
# context_manager.py (partial snippet)
from windows_advanced_control import (
    enable_night_light, disable_night_light, open_night_light_settings,
    enable_dark_mode, disable_dark_mode,click_turn_on_now_opencv,
    set_focus_assist_off, set_focus_assist_priority, set_focus_assist_alarms,
)
from display_control import open_display_settings, open_night_light_settings
from system_monitor import get_system_info
from network_control import open_network_settings, run_speed_test, ping_test
from audio_control_alt import set_volume, volume_up, volume_down, toggle_mute
current_directory = os.path.expanduser("~")

def extract_after_keyword(user_input, keyword):
    if keyword in user_input:
        result = user_input.split(keyword, 1)[1].strip()
        if result.lower().startswith("the "):
            result = result[4:].strip()
        return result
    return None
# In generate_contextual_response function
def generate_contextual_response(user_input, user_id):
    global current_directory
        # Filter out system messages
    if user_input.lower().startswith("listening...") or user_input.lower().startswith("you said:"):
        return "How can I assist you?"
    file_response = handle_file_commands(user_input)
    if file_response is not None:
        update_conversation_history(user_id, user_input, file_response)
        return file_response

    # Limit conversation history to last 3 exchanges
    history = get_conversation_history(user_id)[-3:]
    
    # Build cleaner context
    context = "\n".join(
        [f"User: {item['user'][:100]}\nAssistant: {item['nova'][:200]}" 
         for item in history]
    )
    
    # Create focused prompt
    prompt = f"""Current conversation (most recent last):
{context}
New user input: {user_input}
Provide a helpful, concise response to the user's latest query:"""
    
    nova_response = generate_response(prompt)
    
    # Prevent meta-commentary loops
    if "conversation" in nova_response.lower() and "history" in nova_response.lower():
        nova_response = "How can I assist you with this?"
    
    update_conversation_history(user_id, user_input, nova_response)
    return nova_response
def handle_gmail_commands(user_input):
    user_input_lower = user_input.lower().strip()

    if "open gmail" in user_input_lower:
        return open_gmail()
    if "gmail home" in user_input_lower:
        return gmail_home()
    if "read" in user_input_lower and "mail" in user_input_lower:
        # For example: "read first mail" or "read 1st mail"
        words = user_input_lower.split()
        n = None
        ordinal_map = {
            "first": 1, "1st": 1,
            "second": 2, "2nd": 2,
            "third": 3, "3rd": 3,
            "fourth": 4, "4th": 4,
            "fifth": 5, "5th": 5,
            "sixth" : 6,"6th" : 6,
            "seventh" : 7,"7th":7,
            "eight":8, "8th":8,
            "ninth":9,"9th":9,
            "tenth":10,"10th":10,
        }
        for word in words:
            if word.isdigit():
                n = int(word)
                break
            elif word in ordinal_map:
                n = ordinal_map[word]
                break
        if n is None:
            return "Please specify which mail to read, e.g. 'read first mail'."
        return read_mail(n)
    
    if "close" in user_input_lower and "mail" in user_input_lower:
        return close_gmail()
    if "compose mail" in user_input_lower:
        return handle_gmail_send_command(user_input)
    if "find mail" in user_input_lower:
        # Remove "search mail" prefix to get the query.
        query = user_input_lower.replace("find mail", "").strip()
        if not query:
            query = get_voice_input("Please say what you want to search for in Gmail:")
        return search_mail(query)

    return None

def handle_whatsapp_commands(user_input):
    """
    Handles WhatsApp commands:
    - "open whatsapp"
    - "chat <contact_name>"
    - "send <message> to <contact_name>" or "send message to <phone_number>: <message>"
    """  
    user_input_lower = user_input.lower().strip()

    # Command: "chat <contact_name>"
    if user_input_lower.startswith("chat "):
        contact = user_input_lower[5:].strip()
        if contact:
            return open_chat_with_contact(contact)
        else:
            return "Please specify a contact name for chat."
    # Command: "send message <message>" - sends message to the active chat
    if user_input_lower.startswith("send message"):
        # Remove the "send message" prefix and any extra spaces or colons
        message = user_input[12:].strip(" :")
        if message:
            if send_whatsapp_message(message):
                return f"Sent message: {message}"
            else:
                return "Failed to send message."
        else:
            return "Please provide a message to send."
    # Command: "send <message> to <contact_name>"
    if user_input_lower.startswith("send "):
        # We'll try to split on " to "
        parts = user_input_lower[5:].split(" to ")
        if len(parts) >= 2:
            message = parts[0].strip()
            target = parts[1].strip()
            # If target is numeric (allowing '+' sign), treat as phone number
            if target.replace("+", "").isdigit():
                return send_message_to_number(target, message)
            else:
                return send_message_to_contact(target, message)
        else:
            return "Please specify a message and a target separated by ' to '."
     # Command: "video call" - initiate video call
    if "video call" in user_input_lower:
        return start_video_call()

    # Command: "voice call" - initiate voice call
    if "voice call" in user_input_lower:
        return start_voice_call()

    # Command: "view status" - open the status tab
    if "view status" in user_input_lower:
        return view_status()
    if "close whatsapp" in user_input_lower:
        return close_whatsapp()
    if "silence chart" in user_input_lower or "silence current chat" in user_input_lower:
        return mute_current_chat()

    if "unsilence chart" in user_input_lower or "unsilence current chat" in user_input_lower:
        return unmute_current_chat()
    return None
def handle_file_commands(user_input):
    global current_directory
    user_input_lower = user_input.lower()
    # Check for WhatsApp commands first
    whatsapp_result = handle_whatsapp_commands(user_input)
    if whatsapp_result is not None:
        return whatsapp_result
    gmail_result = handle_gmail_commands(user_input)
    if gmail_result is not None:
        return gmail_result
        # Night Light
    if "enable night light" in user_input_lower:
        return enable_night_light()
    if "disable night light" in user_input_lower:
        return disable_night_light()
    if "display night light" in user_input_lower:
        return open_night_light_settings()

    # Dark Mode
    if "enable dark mode" in user_input_lower:
        return enable_dark_mode()
    if "disable dark mode" in user_input_lower or "enable light mode" in user_input_lower:
        return disable_dark_mode()
    
    # Focus Assist
    if "focus assist off" in user_input_lower:
        return set_focus_assist_off()
    if "focus assist priority" in user_input_lower:
        return set_focus_assist_priority()
    if "focus assist alarms" in user_input_lower:
        return set_focus_assist_alarms()
     # System control commands
    if "display settings" in user_input_lower:
        return open_display_settings()

    if "display night light" in user_input_lower:
        return open_night_light_settings()
    if "turn on now" in user_input_lower:
        return click_turn_on_now_opencv()
    if "network settings" in user_input_lower:
        return open_network_settings()

    if "speed test" in user_input_lower:
        return run_speed_test()

    if "ping" in user_input_lower:
        # For example: "ping 8.8.8.8"
        parts = user_input_lower.split()
        host = parts[1] if len(parts) > 1 else "8.8.8.8"
        return ping_test(host)
    if "set volume to" in user_input_lower:
        try:
            # Expected formats: "set volume to 70%" or "set volume to 70"
            part = user_input_lower.split("set volume to", 1)[1].strip()
            if part.endswith("%"):
                part = part[:-1].strip()
            percent = int(part)
            return set_volume(percent)
        except Exception as e:
            return f"Error processing volume command: {str(e)}"

    if "volume up" in user_input_lower:
        return volume_up()

    if "volume down" in user_input_lower:
        return volume_down()

    if "mute" in user_input_lower or "toggle mute" in user_input_lower:
        return toggle_mute()
    if "system info" in user_input_lower or "system status" in user_input_lower:
        return get_system_info()
    if "brightness" in user_input_lower:
        if "increase" in user_input_lower:
            return increase_brightness()
        elif "decrease" in user_input_lower:
            return decrease_brightness()
        elif "set" in user_input_lower:
            try:
                level = int(user_input_lower.split("set brightness to")[1].strip().replace("%", ""))
                return set_brightness(level)
            except Exception:
                return "Invalid brightness level."
        else:
            return "Please specify 'increase', 'decrease', or 'set brightness to X%'."
    if "restart" in user_input_lower:
        return restart_system()

    if "shutdown" in user_input_lower:
        return shutdown_system()

    if "lock" in user_input_lower:
        return lock_system()
    
    if "open notepad" in user_input_lower:
        return open_application("notepad")

    # if "modify" in user_input_lower and "with notepad" in user_input_lower:
    #     file_name = user_input_lower.replace("modify", "").replace("with notepad", "").strip()
    #     if file_name:
    #         file_path = os.path.join(current_directory, file_name)
    #         return open_file(file_path, "notepad")
    #     return "Specify a file to modify."

    if "type" in user_input_lower:
        text = user_input_lower.replace("type", "").strip()
        pyautogui.write(text)
        return f"Typed: {text}"

    # if "select all" in user_input_lower:
    #     send_keys(["ctrl", "a"])
    #     return "Selected all text."

    # Existing "copy file" command remains unchanged...
    if "copy file" in user_input_lower:
        try:
            command_part = user_input_lower.split("copy file", 1)[1].strip()
            parts = command_part.split(" to ")
            if len(parts) == 2:
                source_file = parts[0].strip()
                dest_file = parts[1].strip()
                source_path = os.path.join(current_directory, source_file)
                dest_path = os.path.join(current_directory, dest_file)
                return copy_file(source_path, dest_path)
            else:
                return "Specify source and destination file names for copying."
        except Exception as e:
            return f"Error processing copy file command: {str(e)}"
    # if "copy" in user_input_lower:
    #     send_keys(["ctrl", "a"])
    #     return "Copied selected text."

    # if "paste" in user_input_lower:
    #     send_keys(["ctrl", "v"])
    #     return "Pasted clipboard content."

    # if "cut" in user_input_lower:
    #     send_keys(["ctrl", "x"])
    #     return "Cut selected text."

    # if "undo" in user_input_lower:
    #     send_keys(["ctrl", "z"])
    #     return "Undid last action."

    # if "redo" in user_input_lower:
    #     send_keys(["ctrl", "y"])
    #     return "Redid last action."

    # if "find" in user_input_lower:
    #     search_text = user_input_lower.replace("find", "").strip()
    #     send_keys(["ctrl", "f"])
    #     time.sleep(1)
    #     pyautogui.write(search_text)
    #     send_keys(["enter"])
    #     return f"Searched for '{search_text}'."

    if "replace" in user_input_lower and "with" in user_input_lower:
        parts = user_input_lower.replace("replace", "").split("with")
        old_text = parts[0].strip()
        new_text = parts[1].strip()
        send_keys(["ctrl", "h"])
        time.sleep(1)
        pyautogui.write(old_text)
        send_keys(["tab"])
        pyautogui.write(new_text)
        send_keys(["enter"])
        return f"Replaced '{old_text}' with '{new_text}'."

    if "save file as" in user_input_lower:
        file_name = user_input_lower.replace("save file as", "").strip()
        send_keys(["ctrl", "shift", "s"])
        time.sleep(1)
        pyautogui.write(file_name)
        send_keys(["enter"])
        return f"Saved file as '{file_name}'."


    if "close notepad" in user_input_lower:
        send_keys(["alt", "f4"])
        return "Closed Notepad."

    if "open file explorer" in user_input_lower:
        response = open_directory(current_directory)
        return response

    if "go out" in user_input_lower:
        new_path = navigate_to_path("go out", current_directory)
        if not new_path.lower().startswith("error"):
            current_directory = new_path
            open_directory(current_directory)
            return f"Moved out to: {current_directory}"
        else:
            return new_path

    if "go to" in user_input_lower:
        path = extract_after_keyword(user_input_lower, "go to")
        if path:
            new_path = navigate_to_path(path, current_directory)
            if not new_path.lower().startswith("error"):
                current_directory = new_path
                open_directory(current_directory)
                return f"Navigated to: {current_directory}"
            else:
                return new_path
        return "Specify a path to navigate to."

    if "list files" in user_input_lower:
        return list_files(current_directory)
    
    if "create folder" in user_input_lower:
        folder_name = extract_after_keyword(user_input_lower, "create folder")
        if folder_name:
            response = create_directory(os.path.join(current_directory, folder_name))
            return response
        return "Specify a folder name."
    
    if "rename" in user_input_lower and " to " in user_input_lower:
        try:
            command_part = user_input_lower.split("rename", 1)[1]
            old_name, new_name = command_part.split(" to ", 1)
            old_name = old_name.strip()
            new_name = new_name.strip()
            if old_name and new_name:
                return rename_file(
                    os.path.join(current_directory, old_name),
                    os.path.join(current_directory, new_name)
                )
            return "Specify old and new names."
        except Exception as e:
            return f"Error processing rename command: {str(e)}"
        # New condition: Return the current path.
    if "give me current path" in user_input_lower or "what is my current path" in user_input_lower:
        return f"Current path is: {current_directory}"
    if "delete" in user_input_lower:
        file_name = extract_after_keyword(user_input_lower, "delete")
        if file_name:
            return delete_file(os.path.join(current_directory, file_name))
        return "Specify a file to delete."
      # New condition: "openfile with <filename> with <application>"
  # New condition: If the command is in the format "<filename> with <application>"
    # and does NOT start with "open", treat it as a file open command.
    if " with " in user_input_lower and not user_input_lower.startswith("open"):
        parts = user_input_lower.split(" with ", 1)
        if len(parts) == 2:
            file_name = parts[0].strip()
            app_name = parts[1].strip()
            if file_name:
                file_path = os.path.join(current_directory, file_name)
                return open_file_with(file_path, app_name)
            else:
                return "Please specify the file name."
      # New condition: "copy <filename>" (without "file" keyword) to copy file content to clipboard
    if user_input_lower.startswith("copy ") and "copy file" not in user_input_lower:
        filename = user_input_lower.split("copy ", 1)[1].strip()
        source_path = os.path.join(current_directory, filename)
        return copy_file(source_path, "")
       # New condition: "paste" command to paste clipboard content
    if user_input_lower.startswith("paste"):
        # Command may be "paste to <destination>"
        parts = user_input_lower.split("to")
        if len(parts) >= 2:
            destination = parts[1].strip()
            destination_path = os.path.join(current_directory, destination)
            return paste_file(destination_path)
        else:
            return "Please specify the destination file name for pasting."

        # Existing condition for "move file" commands
    if "move file" in user_input_lower:
        try:
            command_part = user_input_lower.split("move file", 1)[1].strip()
            parts = command_part.split(" to ")
            if len(parts) == 2:
                source_file = parts[0].strip()
                dest_file = parts[1].strip()
                source_path = os.path.join(current_directory, source_file)
                dest_path = os.path.join(current_directory, dest_file)
                return move_file(source_path, dest_path)
            else:
                return "Specify source and destination for moving file."
        except Exception as e:
            return f"Error processing move file command: {str(e)}"

    # New condition: "move <filename or foldername> to <destination>"
    if user_input_lower.startswith("move ") and "move file" not in user_input_lower:
        try:
            command_part = user_input_lower.split("move ", 1)[1].strip()
            parts = command_part.split(" to ")
            if len(parts) == 2:
                source = parts[0].strip()
                dest = parts[1].strip()
                source_path = os.path.join(current_directory, source)
                dest_path = os.path.join(current_directory, dest)
                return move_file(source_path, dest_path)
            else:
                return "Specify source and destination for moving."
        except Exception as e:
            return f"Error processing move command: {str(e)}"
    if "open" in user_input_lower:
        if "youtube" in user_input_lower:
            return open_application("open youtube")
        elif "browser" in user_input_lower:
            return open_application("open browser")
        elif "vs code" in user_input_lower or "visual studio code" in user_input_lower:
            return open_application("open vs code")
        elif "notepad" in user_input_lower:
            return open_application("open notepad")
        elif "calculator" in user_input_lower:
            return open_application("open calculator")
        elif "file explorer" in user_input_lower:
            return open_application("open file explorer")
        elif "pictures" in user_input_lower:
            return open_application("open pictures")
        elif "documents" in user_input_lower:
            return open_application("open documents")
        elif "downloads" in user_input_lower:
            return open_application("open downloads")
        elif "music" in user_input_lower:
            return open_application("open music")
        else:
            return "Command not recognized."
    return None

def send_keys(keys):
    for key in keys:
        keyboard.press_and_release(key)
        time.sleep(0.1)