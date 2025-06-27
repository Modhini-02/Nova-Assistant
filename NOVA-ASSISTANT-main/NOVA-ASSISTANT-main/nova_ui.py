import sys, os
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QScrollArea, QFrame, QLineEdit, QToolButton,
    QSpacerItem, QSizePolicy
)
from PyQt5.QtGui import QFont
# Import your assistant and voice functions.
from main import NovaAssistant
from voice_handler import listen, speak

os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

##############################################################################
#                               THEME COLORS
##############################################################################
COLORS = {
    "background": "#2D2D2D",
    "sidebar": "#1E1E1E",
    "primary": "#0078D4",
    "text": "#FFFFFF",
    "secondary": "#3C3C3C",
    "success": "#4CAF50",
    "danger": "#F44336",
    "header": "#3C3C3C"
}

LIGHT_COLORS = {
    "background": "#F0F0F0",
    "sidebar": "#D3D3D3",
    "primary": "#0078D4",
    "text": "#000000",
    "secondary": "#E8E8E8",
    "success": "#4CAF50",
    "danger": "#F44336",
    "header": "#D3D3D3"
}

##############################################################################
#                             CHAT BUBBLE
##############################################################################
class ChatBubble(QFrame):
    def __init__(self, text, is_nova=True, parent=None):
        super().__init__(parent)
        self.text = text
        self.is_nova = is_nova
        self.setMinimumWidth(200)
        self.setMaximumWidth(400)
        self.initUI()

    def initUI(self):
        self.label = QLabel(self.text, self)
        self.label.setWordWrap(True)
        self.label.setFont(QFont("Segoe UI", 11))
        self.label.setStyleSheet(f"color: {COLORS['text']};")
        layout = QVBoxLayout(self)
        layout.addWidget(self.label)
        layout.setContentsMargins(15, 10, 15, 10)
        self.setStyleSheet(f"""
            background-color: {COLORS['secondary'] if self.is_nova else COLORS['primary']};
            border-radius: 15px;
            margin: 5px;
        """)

    def update_colors(self):
        self.setStyleSheet(f"""
            background-color: {COLORS['secondary'] if self.is_nova else COLORS['primary']};
            border-radius: 15px;
            margin: 5px;
        """)
        self.label.setStyleSheet(f"color: {COLORS['text']};")

##############################################################################
#                         COMMAND SECTION (Sidebar)
##############################################################################
class CommandSection(QWidget):
    def __init__(self, title, commands, parent=None):
        super().__init__(parent)
        self.title = title
        self.commands = commands
        self.is_expanded = False
        self.initUI()

    def initUI(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        self.header_btn = QPushButton(self.title)
        self.header_btn.setStyleSheet(f"""
            QPushButton {{
                text-align: left;
                padding: 15px;
                font: bold 14px 'Segoe UI';
                color: {COLORS['text']};
                background: {COLORS['sidebar']};
                border: none;
            }}
            QPushButton:hover {{
                background: {COLORS['secondary']};
            }}
        """)
        self.header_btn.clicked.connect(self.toggle_expansion)
        self.layout.addWidget(self.header_btn)
        self.commands_widget = QWidget()
        self.commands_layout = QVBoxLayout(self.commands_widget)
        self.commands_layout.setContentsMargins(25, 0, 0, 0)
        self.commands_layout.setSpacing(5)
        for cmd in self.commands:
            label = QLabel(f"• {cmd}")
            label.setStyleSheet(f"color: {COLORS['text']}; font: 12px 'Segoe UI'; padding: 5px;")
            self.commands_layout.addWidget(label)
        self.layout.addWidget(self.commands_widget)
        self.commands_widget.setVisible(False)

    def toggle_expansion(self):
        self.is_expanded = not self.is_expanded
        self.commands_widget.setVisible(self.is_expanded)

    def update_theme(self):
        self.header_btn.setStyleSheet(f"""
            QPushButton {{
                text-align: left;
                padding: 15px;
                font: bold 14px 'Segoe UI';
                color: {COLORS['text']};
                background: {COLORS['sidebar']};
                border: none;
            }}
            QPushButton:hover {{
                background: {COLORS['secondary']};
            }}
        """)
        for i in range(self.commands_layout.count()):
            label = self.commands_layout.itemAt(i).widget()
            label.setStyleSheet(f"color: {COLORS['text']}; font: 12px 'Segoe UI'; padding: 5px;")

##############################################################################
#                             SIDEBAR (Full Features)
##############################################################################
class Sidebar(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedWidth(300)
        self.setStyleSheet(f"background-color: {COLORS['sidebar']};")
        self.command_sections = []
        self.initUI()

    def initUI(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        self.main_layout.setAlignment(Qt.AlignTop)
        header = QWidget()
        header.setFixedHeight(50)
        self.header = header
        header.setStyleSheet(f"background: {COLORS['header']};")
        self.header_layout = QHBoxLayout(header)
        self.header_layout.setContentsMargins(8, 8, 8, 8)
        self.header_layout.setSpacing(8)
        self.title = QLabel("NOVA")
        self.title.setStyleSheet(f"font: bold 20px 'Segoe UI'; color: {COLORS['text']};")
        self.header_layout.addWidget(self.title)
        self.spacer = QSpacerItem(20, 10, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.header_layout.addItem(self.spacer)
        self.theme_btn = QToolButton()
        self.theme_btn.setText("☀")
        self.theme_btn.setFont(QFont("Segoe UI", 16))
        self.theme_btn.setStyleSheet("border: none; color: " + COLORS['text'] + ";")
        self.header_layout.addWidget(self.theme_btn)
        self.toggle_btn = QToolButton()
        self.toggle_btn.setText("◀")
        self.toggle_btn.setFont(QFont("Segoe UI", 16))
        self.toggle_btn.setStyleSheet("border: none; color: " + COLORS['text'] + ";")
        self.header_layout.addWidget(self.toggle_btn)
        self.main_layout.addWidget(header, 0, Qt.AlignTop)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 15, 0, 15)
        content_layout.setSpacing(10)
        data = {
            "💬 WhatsApp": ["Open WhatsApp", "Close WhatsApp", "Send Message", "View Status", "Voice Call", "Video Call"],
            "📧 Gmail": ["Open Gmail", "Compose Mail", "Find Mail", "Read Mail", "Close Gmail"],
            "🎵 Music": ["Play Music", "Pause Music", "Next Song", "Previous Song", "Volume Up", "Volume Down"],
            "💻 VS Code": ["Open VS Code", "Create File", "Write Code", "Run File", "Close VS Code"],
            "🖥️ System": ["Volume Control", "Brightness", "Display Settings", "Night Light", "Dark Mode"],
            "📁 File Management": ["Open File", "Create Folder", "Delete File", "Rename File", "Move File"]
        }
        for title, commands in data.items():
            section = CommandSection(title, commands)
            self.command_sections.append(section)
            content_layout.addWidget(section)
        content_layout.addStretch()
        scroll.setWidget(content)
        self.scroll = scroll
        self.main_layout.addWidget(scroll)

    def toggle_theme(self):
        global COLORS
        if COLORS["background"] == "#2D2D2D":
            COLORS = LIGHT_COLORS
        else:
            COLORS = {
                "background": "#2D2D2D",
                "sidebar": "#1E1E1E",
                "primary": "#0078D4",
                "text": "#FFFFFF",
                "secondary": "#3C3C3C",
                "success": "#4CAF50",
                "danger": "#F44336",
                "header": "#3C3C3C"
            }
        self.update_ui()

    def update_ui(self):
        self.setStyleSheet(f"background-color: {COLORS['sidebar']};")
        self.header.setStyleSheet(f"background: {COLORS['header']};")
        self.title.setStyleSheet(f"font: bold 20px 'Segoe UI'; color: {COLORS['text']};")
        self.theme_btn.setStyleSheet("border: none; color: " + COLORS['text'] + ";")
        self.toggle_btn.setStyleSheet("border: none; color: " + COLORS['text'] + ";")
        for section in self.command_sections:
            section.update_theme()

##############################################################################
#              VOICE MODE THREAD (Using main2.py's main)
##############################################################################
class VoiceMainThread(QThread):
    logSignal = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

    def run(self):
        # Import and run main() from main2.py
        from main2 import main
        main(self.log)

    def log(self, message):
        self.logSignal.emit(message)

##############################################################################
#                             MAIN WINDOW
##############################################################################
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Nova Assistant")
        self.resize(1400, 800)
        self.assistant = NovaAssistant()
        self.voice_mode_thread = None
        self.voice_mode_active = False
        self.initUI()
        self.setupConnections()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Sidebar with full features.
        self.sidebar = Sidebar()
        main_layout.addWidget(self.sidebar)

        # Content area.
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)

        # Chat area.
        self.chat_area = QScrollArea()
        self.chat_area.setWidgetResizable(True)
        self.chat_area.setStyleSheet("background: transparent; border: none;")
        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setAlignment(Qt.AlignTop)
        self.chat_area.setWidget(self.chat_container)
        content_layout.addWidget(self.chat_area)

        # Input area for text commands, including a Voice button.
        self.input_widget = QWidget()
        input_layout = QHBoxLayout(self.input_widget)
        input_layout.setContentsMargins(0, 10, 0, 0)
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Type your command here...")
        self.input_field.setStyleSheet(f"""
            QLineEdit {{
                padding: 15px;
                border-radius: 25px;
                background: {COLORS['secondary']};
                color: {COLORS['text']};
                border: 2px solid {COLORS['primary']};
                font: 14px 'Segoe UI';
            }}
        """)
        self.send_btn = QPushButton("Send")
        self.send_btn.setFixedSize(100, 50)
        self.send_btn.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['primary']};
                color: {COLORS['text']};
                border-radius: 25px;
                font: bold 14px 'Segoe UI';
            }}
            QPushButton:hover {{
                background: #0062A3;
            }}
        """)
        self.voice_btn = QPushButton("Voice")
        self.voice_btn.setFixedSize(100, 50)
        self.voice_btn.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['primary']};
                color: {COLORS['text']};
                border-radius: 25px;
                font: bold 14px 'Segoe UI';
            }}
            QPushButton:hover {{
                background: #0062A3;
            }}
        """)
        input_layout.addWidget(self.input_field)
        input_layout.addWidget(self.send_btn)
        input_layout.addWidget(self.voice_btn)
        content_layout.addWidget(self.input_widget)

        # Voice mode panel (initially hidden).
        self.voice_mode_panel = QWidget()
        voice_layout = QVBoxLayout(self.voice_mode_panel)
        voice_layout.setContentsMargins(0, 10, 0, 10)
        self.voice_status_label = QLabel("Voice Mode Active. Nova is listening...")
        self.voice_status_label.setFont(QFont("Segoe UI", 16))
        self.voice_status_label.setStyleSheet(f"color: {COLORS['text']};")
        voice_layout.addWidget(self.voice_status_label, alignment=Qt.AlignCenter)
        controls_layout = QHBoxLayout()
        self.start_voice_btn = QPushButton("Start")
        self.start_voice_btn.setFixedSize(80, 40)
        self.start_voice_btn.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['primary']};
                color: {COLORS['text']};
                border-radius: 15px;
            }}
        """)
        self.close_voice_btn = QPushButton("Close")
        self.close_voice_btn.setFixedSize(80, 40)
        self.close_voice_btn.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['primary']};
                color: {COLORS['text']};
                border-radius: 15px;
            }}
        """)
        controls_layout.addWidget(self.start_voice_btn)
        controls_layout.addWidget(self.close_voice_btn)
        voice_layout.addLayout(controls_layout)
        self.voice_mode_panel.setVisible(False)
        content_layout.addWidget(self.voice_mode_panel)

        main_layout.addWidget(content_widget, stretch=1)
        self.setStyleSheet(f"background-color: {COLORS['background']}; color: {COLORS['text']};")

    def setupConnections(self):
        self.send_btn.clicked.connect(self.process_text_input)
        self.input_field.returnPressed.connect(self.process_text_input)
        self.voice_btn.clicked.connect(self.activate_voice_mode)
        self.start_voice_btn.clicked.connect(self.start_voice_thread)
        self.close_voice_btn.clicked.connect(self.deactivate_voice_mode)
        # Also connect sidebar buttons if needed (theme toggle, etc.)
        self.sidebar.theme_btn.clicked.connect(self.sidebar.toggle_theme)
        self.sidebar.toggle_btn.clicked.connect(lambda: None)  # Placeholder for sidebar toggle action

    def add_message(self, text, is_nova=True):
        bubble = ChatBubble(text, is_nova)
        self.chat_layout.addWidget(bubble, alignment=Qt.AlignLeft if is_nova else Qt.AlignRight)
        self.chat_area.verticalScrollBar().setValue(self.chat_area.verticalScrollBar().maximum())

    def add_log(self, message):
        self.add_message(message, is_nova=True)

    def process_text_input(self):
        if self.voice_mode_active:
            return
        text = self.input_field.text().strip()
        if text:
            self.add_message(text, is_nova=False)
            self.input_field.clear()
            response = self.assistant.process_command(text, self.add_log)
            if response:
                self.add_message(response, is_nova=True)

    def activate_voice_mode(self):
        self.voice_mode_active = True
        self.input_widget.setVisible(False)
        self.voice_mode_panel.setVisible(True)

    def start_voice_thread(self):
        if self.voice_mode_thread is None or not self.voice_mode_thread.isRunning():
            self.voice_mode_thread = VoiceMainThread()
            self.voice_mode_thread.logSignal.connect(self.add_log)
            self.voice_mode_thread.start()

    def deactivate_voice_mode(self):
        if self.voice_mode_thread is not None:
            self.voice_mode_thread.terminate()  # Force-stop the thread
            self.voice_mode_thread.wait()
            self.voice_mode_thread = None
        self.voice_mode_active = False
        self.voice_mode_panel.setVisible(False)
        self.input_widget.setVisible(True)

##############################################################################
#                                 RUN
##############################################################################
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
