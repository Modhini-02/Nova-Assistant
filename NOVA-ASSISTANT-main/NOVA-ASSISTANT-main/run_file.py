import os
import subprocess

VS_CODE_FOLDER = "D:/Nova-testing"

def run_specific_file(filename, file_type):
    file_extensions = {
        "python": ".py",
        "c": ".c",
        "c++": ".cpp",
        "java": ".java",
        "html": ".html",
        "javascript": ".js",
        "css": ".css",
    }

    file_type = file_type.lower().strip()  # Convert to lowercase to avoid mismatch

    if file_type not in file_extensions:
        return f"❌ Unsupported file type: {file_type}"

    extension = file_extensions[file_type]
    file_path = os.path.join(VS_CODE_FOLDER, f"{filename}{extension}")

    if not os.path.exists(file_path):
        return f"❌ File '{filename}{extension}' not found in {VS_CODE_FOLDER}."

    run_commands = {
        ".py": f"python \"{file_path}\"",  # Python
        ".c": f"gcc \"{file_path}\" -o \"{file_path[:-2]}\" && \"{file_path[:-2]}\"",  # C
        ".cpp": f"g++ \"{file_path}\" -o \"{file_path[:-4]}\" && \"{file_path[:-4]}\"",  # C++
        ".java": f"javac \"{file_path}\" && java -cp \"{VS_CODE_FOLDER}\" {filename}",  # Java
        ".html": f"start \"\" \"{file_path}\"",  # Open in browser
        ".js": f"node \"{file_path}\"",  # JavaScript (Node.js)
        ".css": f"start \"\" \"{file_path}\"",  # Open CSS in default browser
    }

    print(f"🚀 Running {filename}{extension} ...")

    try:
        # ✅ Capture output and errors
        process = subprocess.Popen(run_commands[extension], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        stdout, stderr = process.communicate()  # Wait for execution

        if process.returncode == 0:
            return f"✅ Execution Output:\n{stdout.strip()}" if stdout else "✅ Execution completed successfully."
        else:
            return f"❌ Error:\n{stderr.strip()}" if stderr else "❌ Execution failed with an unknown error."

    except Exception as e:
        return f"❌ Error running file: {e}"







































# import os
# import subprocess

# VS_CODE_FOLDER = "D:/Testing"


# def run_specific_file(filename, file_type):
#     file_extensions = {
#         "python": ".py",
#         "c": ".c",
#         "c++": ".cpp",
#         "java": ".java",
#         "html": ".html",
#         "javascript": ".js",
#         "css": ".css",
#     }

#     file_type = file_type.lower().strip()  # Convert to lowercase to avoid mismatch

#     if file_type not in file_extensions:
#         print(f" Unsupported file type: {file_type}")
#         return

#     extension = file_extensions[file_type]
#     file_path = os.path.join(VS_CODE_FOLDER, f"{filename}{extension}")

#     if not os.path.exists(file_path):
#         print(f" File '{filename}{extension}' not found in {VS_CODE_FOLDER}.")
#         return

#     run_commands = {
#         ".py": f"python \"{file_path}\"",  # Python
#         ".c": f"gcc \"{file_path}\" -o \"{file_path[:-2]}\" && \"{file_path[:-2]}\"",  # C
#         ".cpp": f"g++ \"{file_path}\" -o \"{file_path[:-4]}\" && \"{file_path[:-4]}\"",  # C++
#         ".java": f"javac \"{file_path}\" && java -cp \"{VS_CODE_FOLDER}\" {filename}",  # Java
#         ".html": f"start \"\" \"{file_path}\"",  # Open in browser
#         ".js": f"node \"{file_path}\"",  # JavaScript (Node.js)
#         ".css": f"start \"\" \"{file_path}\"",  # Open CSS in default browser
#     }

#     print(f"🚀 Running {filename}{extension} ...")
#     response=subprocess.run(run_commands[extension], shell=True)
#     return response
