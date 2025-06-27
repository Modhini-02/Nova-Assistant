import os
import subprocess
import google.generativeai as genai  

# Configure Gemini API
genai.configure(api_key="AIzaSyBHkBd1rszhGlwsumjUhRsiD13jlj2KpTU")

# Load Gemini model
model = genai.GenerativeModel("gemini-2.0-flash")

# Path where VS Code files are stored
VS_CODE_FOLDER = "D:/Nova-testing"

# Function to find the latest file of a specific type
def get_latest_file(extension):
    files = [
        f for f in os.listdir(VS_CODE_FOLDER) if f.endswith(extension) and os.path.isfile(os.path.join(VS_CODE_FOLDER, f))
    ]
    if files:
        latest_file = max(files, key=lambda f: os.path.getctime(os.path.join(VS_CODE_FOLDER, f)))
        return os.path.join(VS_CODE_FOLDER, latest_file)
    return None  # No existing file found

# Function to generate code from Gemini (Forces the correct language)
def get_code_from_gemini(code_description, language):
    prompt = f"Write the {code_description} program in {language}. Provide only the code without explanation."
    response = model.generate_content(prompt)  
    return response.text.strip()

# Function to get or create a file for the specified language
def get_or_create_file(language):
    file_extensions = {
        "python": ".py",
        "c": ".c",
        "c pls plus": ".cpp",
        "java": ".java",
        "html": ".html",
        "javascript": ".js",
        "css": ".css",
    }

    extension = file_extensions.get(language.lower())
    if not extension:
        print(f"❌ Unsupported language: {language}")
        return None

    latest_file = get_latest_file(extension)  # Try to find an existing file
    if latest_file:
        return latest_file  #  Use the latest file if it exists

    # If no file exists, create a new one
    file_path = os.path.join(VS_CODE_FOLDER, f"new{extension}")
    with open(file_path, "w", encoding="utf-8") as file:
        file.write("")  # Create an empty file
    print(f" Created new file: {file_path}")
    return file_path

#  Function to write code into a file (Creates file if needed)
def write_code_in_vscode(code_description, language):
    file_path = get_or_create_file(language)
    if not file_path:
        return

    code = get_code_from_gemini(code_description, language)  # Now explicitly asks for the correct language
    if not code:
        print(" No code generated from Gemini.")
        return

    with open(file_path, "w", encoding="utf-8") as file:
        file.write(code)

    print(f" Code written into {file_path}")
    os.system(f'code "{file_path}"')
