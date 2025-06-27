import tkinter as tk
from tkinter import scrolledtext
from main2 import main
import threading
import sys

class NovaApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Nova Assistant")
        self.root.geometry("600x400")
        self.root.configure(bg="#2E3440")

        # Title Label
        self.title_label = tk.Label(root, text="Nova Assistant", font=("Arial", 24), fg="#88C0D0", bg="#2E3440")
        self.title_label.pack(pady=10)

        # Command Display Area
        self.command_display = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=15, font=("Arial", 12), fg="#D8DEE9", bg="#3B4252")
        self.command_display.pack(pady=10, padx=10)

        # Start Button
        self.start_button = tk.Button(root, text="Start Nova", command=self.start_nova, font=("Arial", 14), fg="#2E3440", bg="#88C0D0")
        self.start_button.pack(pady=10)

        # Exit Button
        self.exit_button = tk.Button(root, text="Exit", command=self.quit_app, font=("Arial", 14), fg="#2E3440", bg="#BF616A")
        self.exit_button.pack(pady=10)

    def start_nova(self):
        self.command_display.insert(tk.END, "Nova is listening...\n")
        threading.Thread(target=self.run_nova, daemon=True).start()

    def run_nova(self):
        try:
            main(self.log_command)
        except Exception as e:
            self.log_command(f"Error: {str(e)}")

    def log_command(self, text):
        self.command_display.insert(tk.END, f"{text}\n")
        self.command_display.yview(tk.END)

    def quit_app(self):
        self.root.destroy()
 
if __name__ == "__main__":
    root = tk.Tk()
    app = NovaApp(root)
    root.mainloop()