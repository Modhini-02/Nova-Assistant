import os
import subprocess
import pyautogui
import time
from ai_integration import ai_response
import ai_integration


VS_CODE_FOLDER = "D:/Nova-testing"

def create_code_file(command):
    global VS_CODE_FOLDER  

    if VS_CODE_FOLDER is None:
        print("No folder detected in VS Code. Please open a folder first.")
        return

    file_types = {
        "python": "py",
        "javascript": "js",
        "html": "html",
        "css": "css",
        "java": "java",
        "c": "c",
        "cpp": "cpp"
    }

    for key, ext in file_types.items():
        if key in command:
            file_name = f"new.{ext}"  # Default file name
            file_path = os.path.join(VS_CODE_FOLDER, file_name)

            with open(file_path, "w") as f:
                f.write("")  # Create an empty file

            os.system(f"code {file_path}")  # Open file in VS Code
            print(f"Created and opened {file_name} in VS Code.")
            return

    print("File type not recognized. Please specify a valid coding file type.")

# def set_vs_code_folder(folder_path):
#     global VS_CODE_FOLDER
#     VS_CODE_FOLDER = folder_path
#     print(f"VS Code folder set to: {folder_path}")

# Function to run the Python file


