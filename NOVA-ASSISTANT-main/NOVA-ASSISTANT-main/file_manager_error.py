import os
import shutil
import subprocess
import re

# Global clipboard for storing file content
CLIPBOARD = None

# Mapping common app names to correct executables
APP_ALIASES = {
    "vs code": "code",  # VS Code CLI
    "notepad++": "notepad++.exe",
    "word": "winword.exe",
    "excel": "excel.exe"
}

# Global variable to store the current directory path.
# It defaults to the user's home directory.
CURRENT_DIRECTORY = os.path.expanduser("~")

def resolve_path(path):
    """Resolve a path relative to CURRENT_DIRECTORY."""
    global CURRENT_DIRECTORY
    if not os.path.isabs(path):
        return os.path.join(CURRENT_DIRECTORY, path)
    return os.path.abspath(path)

def open_directory(path):
    path = resolve_path(path)
    try:
        if "file explorer" in path.lower():
            subprocess.Popen("explorer.exe")
            return "Opened File Explorer."
        
        system_folders = {
            "documents": os.path.expanduser("~/Documents"),
            "downloads": os.path.expanduser("~/Downloads"),
            "desktop": os.path.expanduser("~/Desktop"),
        }
        if path.lower() in system_folders:
            os.startfile(system_folders[path.lower()])
            return f"Opened {path} folder."
        
        if os.path.exists(path):
            os.startfile(path)
            return f"Opened: {path}"
        else:
            return f"Error: Path '{path}' not found."
    except Exception as e:
        return f"Error: {str(e)}"

def create_directory(path):
    path = resolve_path(path)
    try:
        os.makedirs(path, exist_ok=True)
        return f"Created directory: {path}"
    except Exception as e:
        return f"Error: {str(e)}"

def rename_file(old_name, new_name):
    old_path = resolve_path(old_name)
    new_path = resolve_path(new_name)
    try:
        os.rename(old_path, new_path)
        return f"Renamed '{old_path}' to '{new_path}'"
    except Exception as e:
        return f"Error: {str(e)}"

def delete_file(path):
    path = resolve_path(path)
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
            return f"Deleted directory: {path}"
        else:
            os.remove(path)
            return f"Deleted file: {path}"
    except Exception as e:
        return f"Error: {str(e)}"

def list_files(path="."):
    path = resolve_path(path)
    try:
        items = os.listdir(path)
        dirs = [item for item in items if os.path.isdir(os.path.join(path, item))]
        files = [item for item in items if os.path.isfile(os.path.join(path, item))]
        result = f"Directories: {', '.join(dirs)}\nFiles: {', '.join(files)}" if items else "Empty directory."
        return result
    except Exception as e:
        return f"Error: {str(e)}"

def navigate_to_path(path, _=None):
    global CURRENT_DIRECTORY
    path = re.sub(r"^go to\s+", "", path, flags=re.IGNORECASE).strip()
    
    drive_match = re.fullmatch(r"([A-Za-z])\s*drive", path, re.IGNORECASE)
    if drive_match:
        drive_letter = drive_match.group(1).upper()
        new_path = f"{drive_letter}:\\"
        if os.path.exists(new_path):
            CURRENT_DIRECTORY = new_path
            return new_path
        else:
            return f"Error: Drive {drive_letter}: does not exist."
    
    if path.lower() in ["out", "go out"]:
        parent_dir = os.path.dirname(CURRENT_DIRECTORY.rstrip("\\"))
        CURRENT_DIRECTORY = parent_dir if parent_dir else os.path.expanduser("~")
        return CURRENT_DIRECTORY
    
    new_path = os.path.join(CURRENT_DIRECTORY, path) if not os.path.isabs(path) else path
    new_path = os.path.abspath(new_path)
    
    if os.path.exists(new_path) and os.path.isdir(new_path):
        CURRENT_DIRECTORY = new_path
        return CURRENT_DIRECTORY
    else:
        return f"Error: Path '{new_path}' not found or is not a directory."

def open_file(file_path, application=None):
    file_path = resolve_path(file_path)
    try:
        if application:
            application = APP_ALIASES.get(application.lower(), application)
            subprocess.run([application, file_path], check=False, shell=True)
            return f"Opened '{file_path}' with {application}."
        else:
            os.startfile(file_path)
            return f"Opened '{file_path}' with default application."
    except Exception as e:
        return f"Error: {str(e)}"

def open_application(app_name):
    try:
        if app_name.lower() == "notepad":
            subprocess.Popen("notepad.exe")
            return "Opened Notepad."
        return f"Error: Application '{app_name}' not supported yet."
    except Exception as e:
        return f"Error: {str(e)}"

def open_file_with(file_path, app_name):
    file_path = resolve_path(file_path)
    try:
        if not os.path.exists(file_path):
            return f"Error: Path '{file_path}' not found."
        if os.path.isdir(file_path):
            return f"Error: '{file_path}' is a directory. Please specify a file."
        executable = APP_ALIASES.get(app_name.lower(), app_name)
        if executable.lower() == "code":
            os.system(f'{executable} "{file_path}"')
        else:
            subprocess.Popen([executable, file_path], shell=False)
        return f"Opened '{file_path}' with {executable}."
    except Exception as e:
        return f"Error: {str(e)}"

def copy_file(source, destination=""):
    source = resolve_path(source)
    if destination:
        destination = resolve_path(destination)
    try:
        if not destination:
            if os.path.isdir(source):
                return "Error: Cannot copy directory content to clipboard."
            with open(source, "r") as f:
                global CLIPBOARD
                CLIPBOARD = f.read()
            return f"Copied content of '{source}' to clipboard."
        else:
            shutil.copy2(source, destination)
            return f"Copied '{source}' to '{destination}'"
    except Exception as e:
        return f"Error: {str(e)}"

def paste_file(destination):
    destination = resolve_path(destination)
    try:
        if not CLIPBOARD:
            return "Clipboard is empty."
        with open(destination, "w") as f:
            f.write(CLIPBOARD)
        return f"Pasted clipboard content to '{destination}'."
    except Exception as e:
        return f"Error: {str(e)}"

def move_file(source, destination):
    source = resolve_path(source)
    destination = resolve_path(destination)
    try:
        shutil.move(source, destination)
        return f"Moved '{source}' to '{destination}'"
    except Exception as e:
        return f"Error: {str(e)}"
