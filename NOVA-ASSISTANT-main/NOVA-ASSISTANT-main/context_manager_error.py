import os
import pyautogui
import keyboard
import time
from database_handler import get_conversation_history, update_conversation_history
from gemini_handler import generate_response
from file_manager_error import (
    open_directory, create_directory, rename_file, delete_file, list_files,
    navigate_to_path, open_file, open_application, open_file_with, copy_file,
    paste_file, move_file, CURRENT_DIRECTORY, resolve_path
)
from system_control import (
    set_brightness, increase_brightness, decrease_brightness,
    restart_system, shutdown_system, lock_system
)
from gmail_control import handle_gmail_send_command, open_gmail, read_mail, close_gmail, search_mail, get_voice_input
from whatsapp_control import (
    open_whatsapp, open_chat_with_contact, send_message_to_contact,
    send_message_to_number, send_whatsapp_message, start_video_call, view_status, start_voice_call,
    mute_current_chat, unmute_current_chat, close_whatsapp
)
from windows_advanced_control import (
    enable_night_light, disable_night_light, open_night_light_settings,
    enable_dark_mode, disable_dark_mode, click_turn_on_now_opencv,
    set_focus_assist_off, set_focus_assist_priority, set_focus_assist_alarms,
)
from display_control import open_display_settings, open_night_light_settings
from system_monitor import get_system_info
from network_control import open_network_settings, run_speed_test, ping_test
from audio_control_alt import set_volume, volume_up, volume_down, toggle_mute

def extract_after_keyword(user_input, keyword):
    if keyword in user_input:
        result = user_input.split(keyword, 1)[1].strip()
        if result.lower().startswith("the "):
            result = result[4:].strip()
        return result
    return None

def generate_contextual_response(user_input, user_id):
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
        return open_gmail()
    if "read" in user_input_lower and "mail" in user_input_lower:
        words = user_input_lower.split()
        n = None
        ordinal_map = {
            "first": 1, "1st": 1,
            "second": 2, "2nd": 2,
            "third": 3, "3rd": 3,
            "fourth": 4, "4th": 4,
            "fifth": 5, "5th": 5,
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
        query = user_input_lower.replace("find mail", "").strip()
        if not query:
            query = get_voice_input("Please say what you want to search for in Gmail:")
        return search_mail(query)

    return None

def handle_whatsapp_commands(user_input):
    user_input_lower = user_input.lower().strip()

    if user_input_lower.startswith("chat "):
        contact = user_input_lower[5:].strip()
        if contact:
            return open_chat_with_contact(contact)
        else:
            return "Please specify a contact name for chat."
    if user_input_lower.startswith("send message"):
        message = user_input[12:].strip(" :")
        if message:
            if send_whatsapp_message(message):
                return f"Sent message: {message}"
            else:
                return "Failed to send message."
        else:
            return "Please provide a message to send."
    if user_input_lower.startswith("send "):
        parts = user_input_lower[5:].split(" to ")
        if len(parts) >= 2:
            message = parts[0].strip()
            target = parts[1].strip()
            if target.replace("+", "").isdigit():
                return send_message_to_number(target, message)
            else:
                return send_message_to_contact(target, message)
        else:
            return "Please specify a message and a target separated by ' to '."
    if "video call" in user_input_lower:
        return start_video_call()
    if "voice call" in user_input_lower:
        return start_voice_call()
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
    user_input_lower = user_input.lower()
    whatsapp_result = handle_whatsapp_commands(user_input)
    if whatsapp_result is not None:
        return whatsapp_result
    gmail_result = handle_gmail_commands(user_input)
    if gmail_result is not None:
        return gmail_result
    if "enable night light" in user_input_lower:
        return enable_night_light()
    if "disable night light" in user_input_lower:
        return disable_night_light()
    if "display night light" in user_input_lower:
        return open_night_light_settings()
    if "enable dark mode" in user_input_lower:
        return enable_dark_mode()
    if "disable dark mode" in user_input_lower or "enable light mode" in user_input_lower:
        return disable_dark_mode()
    if "focus assist off" in user_input_lower:
        return set_focus_assist_off()
    if "focus assist priority" in user_input_lower:
        return set_focus_assist_priority()
    if "focus assist alarms" in user_input_lower:
        return set_focus_assist_alarms()
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
        parts = user_input_lower.split()
        host = parts[1] if len(parts) > 1 else "8.8.8.8"
        return ping_test(host)
    if "set volume to" in user_input_lower:
        try:
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
    if "type" in user_input_lower:
        text = user_input_lower.replace("type", "").strip()
        pyautogui.write(text)
        return f"Typed: {text}"
    if "close notepad" in user_input_lower:
        send_keys(["alt", "f4"])
        return "Closed Notepad."
    if "save file as" in user_input_lower:
        file_name = user_input_lower.replace("save file as", "").strip()
        send_keys(["ctrl", "shift", "s"])
        time.sleep(1)
        pyautogui.write(file_name)
        send_keys(["enter"])
        return f"Saved file as '{file_name}'."
    if "open file explorer" in user_input_lower:
        response = open_directory(CURRENT_DIRECTORY)
        return response
    if "go out" in user_input_lower:
        new_path = navigate_to_path("go out", CURRENT_DIRECTORY)
        if not new_path.lower().startswith("error"):
            open_directory(new_path)
            return f"Moved out to: {new_path}"
        else:
            return new_path
    if "go to" in user_input_lower:
        path = extract_after_keyword(user_input, "go to")
        if path:
            new_path = navigate_to_path(path)
            if not new_path.lower().startswith("error"):
                open_directory(new_path)
                return f"Navigated to: {new_path}"
            else:
                return new_path
        else:
            return "Specify path."
    if "list files" in user_input_lower:
        return list_files(CURRENT_DIRECTORY)
    if "create folder" in user_input_lower:
        folder_name = extract_after_keyword(user_input, "create folder")
        return create_directory(folder_name) if folder_name else "Specify folder name."
    if "create file" in user_input_lower:
        file_name = extract_after_keyword(user_input_lower, "create file")
        if file_name:
            try:
                file_path = resolve_path(file_name)
                with open(file_path, "w") as f:
                    f.write("")
                return f"Created file: {file_path}"
            except Exception as e:
                return f"Error: {str(e)}"
        return "Specify file name."
    if "delete" in user_input_lower:
        target = extract_after_keyword(user_input, "delete")
        return delete_file(target) if target else "Specify target to delete."
    if "rename" in user_input_lower and " to " in user_input_lower:
        parts = user_input.split("rename", 1)[1].split(" to ", 1)
        return rename_file(parts[0].strip(), parts[1].strip()) if len(parts) == 2 else "Invalid format."
    if user_input_lower.startswith("copy "):
        source = user_input.split("copy", 1)[1].strip()
        return copy_file(source)
    if user_input_lower.startswith("paste to "):
        destination = user_input.split("paste to", 1)[1].strip()
        return paste_file(destination)
    if "move" in user_input_lower and " to " in user_input_lower:
        parts = user_input.split("move", 1)[1].split(" to ", 1)
        return move_file(parts[0].strip(), parts[1].strip()) if len(parts) == 2 else "Specify source and destination."
    if "with " in user_input_lower and not user_input_lower.startswith("open"):
        parts = user_input.split(" with ", 1)
        return open_file_with(parts[0].strip(), parts[1].strip()) if len(parts) == 2 else "Specify file and app."
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
    if "give me current path" in user_input_lower or "what is my current path" in user_input_lower:
        return f"Current path is: {CURRENT_DIRECTORY}"
    return None

def send_keys(keys):
    for key in keys:
        keyboard.press_and_release(key)
        time.sleep(0.1)
