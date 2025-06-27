import sys, random
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, pyqtProperty, QPoint, QSize, QEasingCurve
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QScrollArea, QFrame, QSpacerItem, 
                            QSizePolicy, QListWidget, QListWidgetItem, QStackedWidget,
                            QLineEdit, QToolButton)
from PyQt5.QtGui import (QPainter, QColor, QFont, QPalette, QBrush, QLinearGradient, 
                        QPen, QPainterPath, QIcon, QKeySequence, QImage, QMovie)

THEMES = {
    "dark": {
        "background": "#2D2D2D",
        "sidebar": "#1E1E1E",
        "primary": "#4A4A4A",
        "text": "#FFFFFF",
        "secondary": "#3C3C3C",
        "success": "#4CAF50",
        "danger": "#F44336",
        "header": "#333333"
    },
    "light": {
        "background": "#F5F5F5",
        "sidebar": "#E0E0E0",
        "primary": "#D3D3D3",
        "text": "#000000",
        "secondary": "#FFFFFF",
        "success": "#4CAF50",
        "danger": "#F44336",
        "header": "#CCCCCC"
    }
}

class ChatBubble(QFrame):
    def __init__(self, text, is_nova=True, parent=None, theme="dark"):
        super().__init__(parent)
        self.text = text
        self.is_nova = is_nova
        self.theme = theme
        self.setMinimumWidth(200)
        self.setMaximumWidth(400)
        self.initUI()
        self.setupAnimation()

    def initUI(self):
        self.label = QLabel(self.text, self)
        self.label.setWordWrap(True)
        self.label.setStyleSheet(f"color: {THEMES[self.theme]['text']};")
        self.label.setFont(QFont("Segoe UI", 11))
        
        layout = QVBoxLayout(self)
        layout.addWidget(self.label)
        layout.setContentsMargins(15, 10, 15, 10)
        
        self.setStyleSheet(f"""
            background-color: {THEMES[self.theme]['secondary'] if self.is_nova else THEMES[self.theme]['primary']};
            border-radius: 15px;
            margin: 5px;
        """)

    def setupAnimation(self):
        self.anim = QPropertyAnimation(self, b"pos")
        self.anim.setDuration(300)
        self.anim.setEasingCurve(QEasingCurve.OutBack)

    def enterEvent(self, event):
        self.anim.setStartValue(self.pos())
        self.anim.setEndValue(self.pos() + QPoint(0, -5))
        self.anim.start()

    def leaveEvent(self, event):
        self.anim.setStartValue(self.pos())
        self.anim.setEndValue(self.pos() + QPoint(0, 5))
        self.anim.start()

class VoiceVisualizer(QWidget):
    def __init__(self, parent=None, theme="dark"):
        super().__init__(parent)
        self.theme = theme
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)
        self.amplitude = 0
        self.setFixedSize(200, 50)
        self.timer.start(50)
        
        self.gradient = QLinearGradient(0, 0, self.width(), 0)
        self.update_colors()

    def update_colors(self):
        primary = QColor(THEMES[self.theme]['primary'])
        accent = QColor("#00B4FF") if self.theme == "dark" else QColor("#0078D4")
        self.gradient.setColorAt(0, primary)
        self.gradient.setColorAt(1, accent)

    def setAmplitude(self, value):
        self.amplitude = value

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        path = QPainterPath()
        path.moveTo(0, self.height()/2)
        
        for x in range(0, self.width(), 5):
            y = self.height()/2 + random.randint(-10, 10) * self.amplitude
            path.lineTo(x, y)
            
        painter.setPen(QPen(self.gradient, 3))
        painter.drawPath(path)

class CommandSection(QWidget):
    def __init__(self, title, commands, parent=None, theme="dark"):
        super().__init__(parent)
        self.theme = theme
        self.title = title
        self.commands = commands
        self.is_expanded = False
        self.initUI()
        self.setupAnimation()

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
                color: {THEMES[self.theme]['text']};
                background: {THEMES[self.theme]['sidebar']};
                border: none;
            }}
            QPushButton:hover {{
                background: {THEMES[self.theme]['secondary']};
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
            label.setStyleSheet(f"color: {THEMES[self.theme]['text']}; font: 12px 'Segoe UI'; padding: 5px;")
            self.commands_layout.addWidget(label)
        
        self.layout.addWidget(self.commands_widget)
        self.commands_widget.setVisible(False)

    def setupAnimation(self):
        self.animation = QPropertyAnimation(self.commands_widget, b"maximumHeight")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)

    def toggle_expansion(self):
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.animation.setStartValue(0)
            self.animation.setEndValue(self.commands_layout.sizeHint().height())
        else:
            self.animation.setEndValue(0)
            self.animation.setStartValue(self.commands_layout.sizeHint().height())
        self.animation.start()
        self.commands_widget.setVisible(self.is_expanded)

class Sidebar(QWidget):
    def __init__(self, theme):
        super().__init__()
        self.theme = theme
        self.setMinimumWidth(50)
        self.setMaximumWidth(300)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Header
        self.header = QWidget()
        self.header.setStyleSheet(f"background: {THEMES[self.theme]['header']};")
        header_layout = QHBoxLayout(self.header)
        header_layout.setContentsMargins(15, 10, 15, 10)
        
        self.title = QLabel("NOVA")
        self.title.setStyleSheet(f"font: bold 20px 'Segoe UI'; color: {THEMES[self.theme]['text']};")
        
        self.toggle_btn = QToolButton()
        self.toggle_btn.setIcon(QIcon("close_icon.png"))
        self.toggle_btn.setIconSize(QSize(20, 20))
        self.toggle_btn.setStyleSheet("border: none;")
        
        self.theme_toggle = QToolButton()
        self.theme_toggle.setIcon(QIcon("theme_toggle.png"))
        self.theme_toggle.setIconSize(QSize(20, 20))
        self.theme_toggle.setStyleSheet("border: none;")
        
        header_layout.addWidget(self.title)
        header_layout.addWidget(self.theme_toggle)
        header_layout.addWidget(self.toggle_btn)
        layout.addWidget(self.header)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.content = QWidget()
        content_layout = QVBoxLayout(self.content)
        
        command_sections = {
            "💬 WhatsApp": ["Open WhatsApp", "Close WhatsApp", "Send Message", "View Status", "Voice Call", "Video Call"],
            "📧 Gmail": ["Open Gmail", "Compose Mail", "Find Mail", "Read Mail", "Close Gmail"],
            "🎵 Music": ["Play Music", "Pause Music", "Next Song", "Previous Song", "Volume Up", "Volume Down"],
            "💻 VS Code": ["Open VS Code", "Create File", "Write Code", "Run File", "Close VS Code"],
            "🖥️ System": ["Volume Control", "Brightness", "Display Settings", "Night Light", "Dark Mode"],
            "📁 File Management": ["Open File", "Create Folder", "Delete File", "Rename File", "Move File"]
        }
        
        for title, commands in command_sections.items():
            section = CommandSection(title, commands, theme=self.theme)
            content_layout.addWidget(section)
        
        scroll.setWidget(self.content)
        layout.addWidget(scroll)
        self.update_styles()

    def update_styles(self):
        self.setStyleSheet(f"background-color: {THEMES[self.theme]['sidebar']};")
        self.header.setStyleSheet(f"background-color: {THEMES[self.theme]['header']};")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_theme = "dark"
        self.initUI()
        self.setupConnections()
        self.apply_theme()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        self.sidebar = Sidebar(self.current_theme)
        self.sidebar_animation = QPropertyAnimation(self.sidebar, b"minimumWidth")
        self.sidebar_animation.setDuration(300)
        main_layout.addWidget(self.sidebar)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)
        
        self.chat_area = QScrollArea()
        self.chat_area.setWidgetResizable(True)
        self.chat_area.setStyleSheet("background: transparent; border: none;")
        
        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setAlignment(Qt.AlignTop)
        self.chat_layout.setSpacing(15)
        self.chat_area.setWidget(self.chat_container)
        content_layout.addWidget(self.chat_area)
        
        input_widget = QWidget()
        input_layout = QHBoxLayout(input_widget)
        input_layout.setContentsMargins(0, 10, 0, 0)
        
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Type your command here...")
        self.update_input_styles()
        
        self.send_btn = QPushButton("Send")
        self.send_btn.setFixedSize(100, 50)
        self.update_button_styles()
        
        self.voice_btn = QPushButton()
        self.voice_btn.setFixedSize(60, 60)
        self.update_voice_button()
        
        input_layout.addWidget(self.input_field)
        input_layout.addWidget(self.voice_btn)
        input_layout.addWidget(self.send_btn)
        content_layout.addWidget(input_widget)
        
        main_layout.addWidget(content_widget, stretch=1)
        self.apply_theme()

    def setupConnections(self):
        self.sidebar.toggle_btn.clicked.connect(self.toggle_sidebar)
        self.sidebar.theme_toggle.clicked.connect(self.toggle_theme)
        self.input_field.returnPressed.connect(self.process_text_input)
        self.send_btn.clicked.connect(self.process_text_input)
        self.voice_btn.clicked.connect(self.toggle_voice_mode)

    def setupVoiceAnimation(self):
        self.voice_visualizer = VoiceVisualizer(theme=self.current_theme)
        self.voice_animation = QPropertyAnimation(self.voice_visualizer, b"amplitude")
        self.voice_animation.setDuration(1000)
        self.voice_animation.setLoopCount(-1)
        self.voice_animation.setStartValue(0)
        self.voice_animation.setEndValue(1)

    def toggle_sidebar(self):
        if self.sidebar.width() > 100:
            self.sidebar_animation.setStartValue(self.sidebar.width())
            self.sidebar_animation.setEndValue(50)
            self.sidebar.toggle_btn.setIcon(QIcon("open_icon.png"))
        else:
            self.sidebar_animation.setStartValue(self.sidebar.width())
            self.sidebar_animation.setEndValue(300)
            self.sidebar.toggle_btn.setIcon(QIcon("close_icon.png"))
        self.sidebar_animation.start()

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme()

    def apply_theme(self):
        theme = THEMES[self.current_theme]
        # Update main window
        self.setStyleSheet(f"""
            background-color: {theme['background']}; 
            color: {theme['text']};
        """)
        
        # Update sidebar
        self.sidebar.theme = self.current_theme
        self.sidebar.update_styles()
        
        # Update chat bubbles
        for i in range(self.chat_layout.count()):
            widget = self.chat_layout.itemAt(i).widget()
            if isinstance(widget, ChatBubble):
                widget.theme = self.current_theme
                widget.initUI()
        
        # Update input components
        self.update_input_styles()
        self.update_button_styles()
        self.update_voice_button()

    def update_input_styles(self):
        theme = THEMES[self.current_theme]
        self.input_field.setStyleSheet(f"""
            QLineEdit {{
                padding: 15px;
                border-radius: 25px;
                background: {theme['secondary']};
                color: {theme['text']};
                border: 2px solid {theme['primary']};
                font: 14px 'Segoe UI';
            }}
        """)

    def update_button_styles(self):
        theme = THEMES[self.current_theme]
        self.send_btn.setStyleSheet(f"""
            QPushButton {{
                background: {theme['primary']};
                color: {theme['text']};
                border-radius: 25px;
                font: bold 14px 'Segoe UI';
            }}
            QPushButton:hover {{
                background: {theme['secondary']};
            }}
        """)

    def update_voice_button(self):
        theme = THEMES[self.current_theme]
        self.voice_btn.setStyleSheet(f"""
            QPushButton {{
                background: {theme['primary']};
                border-radius: 30px;
                border: none;
            }}
            QPushButton:hover {{
                background: {theme['secondary']};
            }}
        """)

    def toggle_voice_mode(self):
        if not hasattr(self, 'voice_mode_active'):
            self.voice_mode_active = False
        
        self.voice_mode_active = not self.voice_mode_active
        if self.voice_mode_active:
            self.voice_animation.start()
            self.add_message("Voice mode activated. Listening...", True)
        else:
            self.voice_animation.stop()
            self.add_message("Voice mode deactivated.", True)

    def process_text_input(self):
        text = self.input_field.text()
        if text:
            self.add_message(text, False)
            self.input_field.clear()
            QTimer.singleShot(1000, lambda: self.simulate_nova_response(text))

    def simulate_nova_response(self, text):
        self.add_message(f"Processed command: {text}", True)
        self.animate_response_bubble()

    def animate_response_bubble(self):
        last_bubble = self.chat_layout.itemAt(self.chat_layout.count()-1).widget()
        anim = QPropertyAnimation(last_bubble, b"geometry")
        anim.setDuration(500)
        anim.setEasingCurve(QEasingCurve.OutBack)
        anim.setStartValue(last_bubble.geometry().adjusted(-10, -10, 10, 10))
        anim.setEndValue(last_bubble.geometry())
        anim.start()

    def add_message(self, text, is_nova=True):
        bubble = ChatBubble(text, is_nova, theme=self.current_theme)
        self.chat_layout.addWidget(bubble, alignment=Qt.AlignLeft if is_nova else Qt.AlignRight)
        self.chat_area.verticalScrollBar().setValue(self.chat_area.verticalScrollBar().maximum())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())