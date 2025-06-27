import sys, os, random
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, pyqtProperty, QRectF
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QScrollArea, QFrame, QLineEdit, QToolButton,
    QSpacerItem, QSizePolicy
)
from PyQt5.QtGui import QFont, QPainter, QColor, QRadialGradient, QLinearGradient, QPen

# Import your assistant and voice functions.
from main import NovaAssistant
from voice_handler import listen, speak

os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

# Define a base blue color for voice animation.
BASE_BLUE = QColor(135, 206, 250)

##############################################################################
#                              THEME COLORS
##############################################################################
# Dark mode
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
# Light mode
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
#                           CHAT BUBBLE
##############################################################################
class ChatBubble(QFrame):
    def __init__(self, text, is_nova=True, parent=None):
        super().__init__(parent)
        self.text = text
        self.is_nova = is_nova
        self.setMinimumWidth(200)
        # No maximum width so that long messages are shown completely.
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.initUI()

    def initUI(self):
        self.label = QLabel(self.text, self)
        self.label.setWordWrap(True)
        self.label.setFont(QFont("Segoe UI", 11))
        self.label.setStyleSheet(f"color: {COLORS['text']};")
        self.label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout = QVBoxLayout(self)
        layout.addWidget(self.label)
        layout.setContentsMargins(15, 10, 15, 10)
        bg_color = COLORS['secondary'] if self.is_nova else COLORS['primary']
        self.setStyleSheet(f"""
            background-color: {bg_color};
            border-radius: 15px;
            margin: 5px;
        """)

    def update_colors(self):
        bg_color = COLORS['secondary'] if self.is_nova else COLORS['primary']
        self.setStyleSheet(f"""
            background-color: {bg_color};
            border-radius: 15px;
            margin: 5px;
        """)
        self.label.setStyleSheet(f"color: {COLORS['text']};")

##############################################################################
#                COMMAND SECTION & SIDEBAR
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

class Sidebar(QWidget):
    def __init__(self):
        super().__init__()
        self.command_sections = []
        self.initUI()

    def initUI(self):
        # In dark mode, use light sidebar; in light mode, use dark sidebar.
        if COLORS["background"] == "#2D2D2D":
            self.setStyleSheet("background-color: #f5f5f5;")
        else:
            self.setStyleSheet("background-color: #333333;")
        self.setFixedWidth(300)
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
            "💬 WhatsApp": ["Open WhatsApp", "Close WhatsApp", "Send Message to person name","send message message", "View Status", "Voice Call", "Video Call"],
            "📧 Gmail": ["Open Gmail", "Compose Mail", "Find Mail mail", "Read nth Mail", "Close Gmail"],
            "🎵 Music": ["open media player/music player","Play Music", "Pause Music", "Next Song", "Previous Song", "Volume Up", "Volume Down","stop music(to close music player)"],
            "💻 VS Code": ["Open VS Code", "Create File", "Write Code(writes code from online)", "Run File", "Close VS Code"],
            "🖥️ System": ["Volume up","volume down","set volume to %","toggle mute", "Increase Brightness","Decrease Brightness","set brightness to %", "Display Settings", "Display Night Light", "Enable Dark Mode","Disable dark mode","Increase Volume","Decrease Volume","system info","speed test","restart","shutdown"],
            "📁 File Management": ["Open File Explorer","go to directory name","go out from directory name","create file filename", "Create Folder foldername", "Delete Filename/foldername", "Rename Filename/foldername with newfilename/foldername", "Move Filename to foldername ","move foldername to drive name","move foldername to foldername","Copy filename","Paste","open filename with applicationname"],
            "🌐 Browser": ["Open Browser","Search Somthing","Close Browser","search google something","click link linknumbeer"],
            "🔴 Youtube ": ["Open Youtube","Play videoname on youtube","Pause youtube","Play youtube","skip add","like video","dislike video","full screen","next video","search video" , "click video videonumber"]
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
        if COLORS["background"] == "#2D2D2D":
            self.setStyleSheet("background-color: #f5f5f5;")
        else:
            self.setStyleSheet("background-color: #333333;")
        self.header.setStyleSheet(f"background: {COLORS['header']};")
        self.title.setStyleSheet(f"font: bold 20px 'Segoe UI'; color: {COLORS['text']};")
        self.theme_btn.setStyleSheet("border: none; color: " + COLORS['text'] + ";")
        self.toggle_btn.setStyleSheet("border: none; color: " + COLORS['text'] + ";")
        for section in self.command_sections:
            section.update_theme()

##############################################################################
#         VOICE ASSISTANT WIDGET (Animation)
##############################################################################
class VoiceAssistantWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Increase widget size for a larger animation
        self.setFixedSize(250, 250)
        self.voice_state = "idle"  # "idle", "listening", "responding"
        self._scaleFactor = 1.0
        self.ripple_phase = 0
        self.num_bars = 12
        self.bar_values = [0.5 for _ in range(self.num_bars)]
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_animation)
        self.timer.start(150)

    def set_voice_state(self, state: str):
        self.voice_state = state
        self.update()

    def update_animation(self):
        if self.voice_state == "idle":
            self._scaleFactor = random.uniform(1.0, 1.1)
        elif self.voice_state == "listening":
            self._scaleFactor = random.uniform(1.0, 1.15)
            self.ripple_phase = (self.ripple_phase + 15) % 360
        elif self.voice_state == "responding":
            self.bar_values = [random.uniform(0.3, 1.0) for _ in range(self.num_bars)]
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        cx = self.width() / 2.0
        cy = self.height() / 2.0

        if self.voice_state == "idle":
            radius = int(90 * self._scaleFactor)
            gradient = QRadialGradient(cx, cy, radius * 1.4, cx, cy)
            gradient.setColorAt(0.0, QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 220))
            gradient.setColorAt(1.0, QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 0))
            painter.setBrush(gradient)
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(int(cx - radius), int(cy - radius), int(radius * 2), int(radius * 2))

        elif self.voice_state == "listening":
            main_radius = int(80 * self._scaleFactor)
            gradient = QRadialGradient(cx, cy, main_radius, cx, cy)
            gradient.setColorAt(0.0, QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 240))
            gradient.setColorAt(1.0, QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 150))
            painter.setBrush(gradient)
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(int(cx - main_radius), int(cy - main_radius), int(main_radius * 2), int(main_radius * 2))
            ripple_count = 2
            for i in range(1, ripple_count + 1):
                ripple_offset = ((self.ripple_phase + i * 30) % 360) / 360.0 * 20
                ripple_radius = int(main_radius + ripple_offset + i * 5)
                opacity = max(0, 150 - i * 60)
                pen = QPen(QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), opacity))
                pen.setWidth(3)
                painter.setPen(pen)
                painter.setBrush(Qt.NoBrush)
                painter.drawEllipse(int(cx - ripple_radius), int(cy - ripple_radius), int(ripple_radius * 2), int(ripple_radius * 2))
            small_radius = int(22 * self._scaleFactor)
            offset_x = cx - main_radius * 0.6
            offset_y = cy + main_radius * 0.6
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 240))
            painter.drawEllipse(int(offset_x - small_radius), int(offset_y - small_radius), int(small_radius * 2), int(small_radius * 2))

        elif self.voice_state == "responding":
            spacing = 4
            total_width = self.width()
            bar_width = (total_width - (self.num_bars + 1) * spacing) / self.num_bars
            max_bar_height = self.height() * 0.8
            for i in range(self.num_bars):
                value = self.bar_values[i]
                bar_height = value * max_bar_height
                x = spacing + i * (bar_width + spacing)
                y = (self.height() - bar_height) / 2.0
                grad = QLinearGradient(x, y, x, y + bar_height)
                grad.setColorAt(0.0, QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 220))
                grad.setColorAt(1.0, QColor(BASE_BLUE.red(), BASE_BLUE.green(), BASE_BLUE.blue(), 150))
                painter.setBrush(grad)
                painter.setPen(Qt.NoPen)
                painter.drawRect(int(x), int(y), int(bar_width), int(bar_height))

##############################################################################
#   VOICE WORKER THREAD (Infinite Loop like nova_app.py)
##############################################################################
class VoiceWorker(QThread):
    userSaidSignal = pyqtSignal(str)
    novaSaidSignal = pyqtSignal(str)
    stateSignal = pyqtSignal(str)
    finishedSignal = pyqtSignal()

    def __init__(self, assistant: NovaAssistant):
        super().__init__()
        self.assistant = assistant
        self.running = True

    def run(self):
        while self.running:
            self.stateSignal.emit("listening")
            command = listen()  # Blocking call
            if not command or command.strip().lower() in ["listening...", "sorry, i didn't catch that.", ""]:
                self.novaSaidSignal.emit("No speech detected")
                continue
            if any(x in command.lower() for x in ["exit", "quit", "stop"]):
                self.novaSaidSignal.emit("Goodbye!")
                speak("Goodbye!")
                break
            self.stateSignal.emit("responding")
            self.userSaidSignal.emit(command)
            response = self.assistant.process_command(command, lambda x: None, speak_response=False)
            if response:
                self.novaSaidSignal.emit(response)
                speak(response)
            self.stateSignal.emit("idle")
        self.running = False
        self.finishedSignal.emit()

    def stop(self):
        self.running = False

##############################################################################
#   MAIN WINDOW
##############################################################################
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Nova Assistant")
        self.resize(1400, 800)
        self.assistant = NovaAssistant()
        self.voice_thread = None
        self.initUI()
        self.setupConnections()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.sidebar = Sidebar()
        main_layout.addWidget(self.sidebar)

        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(20)

        # Left: Chat and text input
        self.text_mode_container = QWidget()
        text_mode_layout = QVBoxLayout(self.text_mode_container)
        text_mode_layout.setContentsMargins(0, 0, 0, 0)
        text_mode_layout.setSpacing(10)

        self.chat_area = QScrollArea()
        self.chat_area.setWidgetResizable(True)
        self.chat_area.setStyleSheet("background: transparent; border: none;")
        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setAlignment(Qt.AlignTop)
        self.chat_area.setWidget(self.chat_container)
        text_mode_layout.addWidget(self.chat_area)

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
        text_mode_layout.addWidget(self.input_widget)

        content_layout.addWidget(self.text_mode_container, stretch=3)

        # Right: Voice mode panel (only Start and Close buttons remain)
        self.voice_mode_panel = QWidget()
        self.voice_mode_panel.setFixedWidth(300)
        voice_layout = QVBoxLayout(self.voice_mode_panel)
        voice_layout.setContentsMargins(10, 10, 10, 10)
        voice_layout.setSpacing(10)

        anim_layout = QHBoxLayout()
        anim_layout.setContentsMargins(0, 0, 0, 0)
        anim_layout.setSpacing(0)
        anim_layout.addStretch()
        self.voice_assistant_widget = VoiceAssistantWidget()
        anim_layout.addWidget(self.voice_assistant_widget)
        voice_layout.addLayout(anim_layout)

        controls_layout = QHBoxLayout()
        self.start_btn = QPushButton("Start")
        self.start_btn.setFixedSize(80, 40)
        self.start_btn.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['primary']};
                color: {COLORS['text']};
                border-radius: 15px;
            }}
            QPushButton:hover {{
                background: #0062A3;
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
            QPushButton:hover {{
                background: #0062A3;
            }}
        """)
        controls_layout.addWidget(self.start_btn)
        controls_layout.addWidget(self.close_voice_btn)
        voice_layout.addLayout(controls_layout)

        self.voice_mode_panel.setVisible(False)
        content_layout.addWidget(self.voice_mode_panel, stretch=1)

        main_layout.addWidget(content_widget, stretch=1)
        self.setStyleSheet(f"background-color: {COLORS['background']}; color: {COLORS['text']};")

    def setupConnections(self):
        self.sidebar.toggle_btn.clicked.connect(self.toggle_sidebar)
        self.sidebar.theme_btn.clicked.connect(self.toggle_theme)
        self.input_field.returnPressed.connect(self.process_text_input)
        self.send_btn.clicked.connect(self.process_text_input)
        self.voice_btn.clicked.connect(self.enter_voice_mode)
        self.close_voice_btn.clicked.connect(self.exit_voice_mode)
        self.start_btn.clicked.connect(self.start_voice_listening)

    def add_message(self, text, is_nova=True):
        bubble = ChatBubble(text, is_nova)
        self.chat_layout.addWidget(bubble, alignment=Qt.AlignLeft if is_nova else Qt.AlignRight)
        self.chat_area.verticalScrollBar().setValue(self.chat_area.verticalScrollBar().maximum())

    def process_text_input(self):
        if self.voice_mode_panel.isVisible():
            return
        text = self.input_field.text().strip()
        if text:
            self.add_message(text, is_nova=False)
            self.input_field.clear()
            response = self.assistant.process_command(text, lambda x: self.add_message(x, is_nova=True))
            if response:
                self.add_message(response, is_nova=True)

    def toggle_sidebar(self):
        current_width = self.sidebar.width()
        if current_width > 140:
            new_width = 140
            self.sidebar.setFixedWidth(new_width)
            if self.sidebar.header_layout.indexOf(self.sidebar.spacer) != -1:
                self.sidebar.header_layout.removeItem(self.sidebar.spacer)
            self.sidebar.theme_btn.setVisible(False)
            self.sidebar.toggle_btn.setText("▶")
            self.sidebar.scroll.setVisible(False)
        else:
            new_width = 300
            self.sidebar.setFixedWidth(new_width)
            if self.sidebar.header_layout.indexOf(self.sidebar.spacer) == -1:
                self.sidebar.header_layout.insertItem(1, self.sidebar.spacer)
            self.sidebar.theme_btn.setVisible(True)
            self.sidebar.toggle_btn.setText("◀")
            self.sidebar.scroll.setVisible(True)

    def toggle_theme(self):
        self.sidebar.toggle_theme()
        self.setStyleSheet(f"background-color: {COLORS['background']}; color: {COLORS['text']};")
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
        for btn in [self.start_btn, self.close_voice_btn]:
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: {COLORS['primary']};
                    color: {COLORS['text']};
                    border-radius: 15px;
                }}
            """)
        for i in range(self.chat_layout.count()):
            widget = self.chat_layout.itemAt(i).widget()
            if isinstance(widget, ChatBubble):
                widget.update_colors()

    # ----- Voice Mode -----
    def enter_voice_mode(self):
        self.input_widget.setVisible(False)
        self.voice_mode_panel.setVisible(True)
        self.voice_assistant_widget.set_voice_state("idle")

    def exit_voice_mode(self):
        if self.voice_thread:
            self.voice_thread.stop()
            self.voice_thread.wait()
            self.voice_thread = None
        self.input_widget.setVisible(True)
        self.voice_mode_panel.setVisible(False)
        self.voice_assistant_widget.set_voice_state("idle")

    def start_voice_listening(self):
        if self.voice_thread:
            return
        self.voice_thread = VoiceWorker(self.assistant)
        self.voice_thread.userSaidSignal.connect(lambda text: self.add_message(text, is_nova=False))
        self.voice_thread.novaSaidSignal.connect(lambda text: self.add_message(text, is_nova=True))
        self.voice_thread.stateSignal.connect(self.voice_assistant_widget.set_voice_state)
        self.voice_thread.finishedSignal.connect(self.on_voice_finished)
        self.voice_thread.start()

    def on_voice_finished(self):
        self.voice_thread = None
        self.voice_assistant_widget.set_voice_state("idle")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
