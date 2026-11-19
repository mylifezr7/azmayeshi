"""
Chat Widget and Message Bubbles for Persian Excel Assistant.
"""
import time
from typing import Callable, Optional
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextBrowser,
    QPushButton, QFrame, QScrollArea, QSizePolicy
)


class MessageBubble(QFrame):
    """Single chat bubble widget with RTL text and custom styling."""

    def __init__(self, sender: str, text: str, is_user: bool = False,
                 is_confirmation: bool = False, on_confirm: Optional[Callable] = None,
                 on_cancel: Optional[Callable] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.is_user = is_user
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(6)

        # Header: sender name + time
        time_str = time.strftime("%H:%M")
        header_text = f"👤 {sender} ({time_str})" if is_user else f"🤖 {sender} ({time_str})"
        header_lbl = QLabel(header_text)
        header_lbl.setStyleSheet("font-size: 11px; color: #94A3B8; font-weight: bold;")
        layout.addWidget(header_lbl)

        # Body text
        body_lbl = QLabel(text)
        body_lbl.setWordWrap(True)
        body_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        body_lbl.setStyleSheet("font-size: 13px; line-height: 1.4; color: #FFFFFF;")
        layout.addWidget(body_lbl)

        # Set bubble styles
        if is_confirmation:
            self.setStyleSheet("""
                QFrame {
                    background-color: #3D2D1E;
                    border: 1px solid #D97706;
                    border-radius: 12px;
                }
            """)
            btn_box = QHBoxLayout()
            btn_box.setSpacing(10)
            btn_confirm = QPushButton("✅ تأیید و اجرا")
            btn_confirm.setStyleSheet("background-color: #107C41; color: white; padding: 6px 14px; border-radius: 6px;")
            if on_confirm:
                btn_confirm.clicked.connect(on_confirm)

            btn_cancel = QPushButton("❌ انصراف")
            btn_cancel.setStyleSheet("background-color: #7F1D1D; color: white; padding: 6px 14px; border-radius: 6px;")
            if on_cancel:
                btn_cancel.clicked.connect(on_cancel)

            btn_box.addWidget(btn_confirm)
            btn_box.addWidget(btn_cancel)
            btn_box.addStretch()
            layout.addLayout(btn_box)

        elif is_user:
            self.setStyleSheet("""
                QFrame {
                    background-color: #1E293B;
                    border: 1px solid #334155;
                    border-radius: 12px;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame {
                    background-color: #132E20;
                    border: 1px solid #107C41;
                    border-radius: 12px;
                }
            """)


class ChatWidget(QWidget):
    """Main scrollable chat feed."""

    confirm_action_signal = Signal(object)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setLayoutDirection(Qt.RightToLeft)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Scroll Area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(16, 16, 16, 16)
        self.container_layout.setSpacing(12)
        self.container_layout.addStretch()

        self.scroll_area.setWidget(self.container)
        main_layout.addWidget(self.scroll_area)

    def add_user_message(self, text: str):
        bubble = MessageBubble(sender="شما", text=text, is_user=True)
        self._append_bubble(bubble)

    def add_assistant_message(self, text: str):
        bubble = MessageBubble(sender="دستیار اکسل", text=text, is_user=False)
        self._append_bubble(bubble)

    def add_confirmation_request(self, prompt_text: str, on_confirm: Callable, on_cancel: Callable):
        bubble = MessageBubble(
            sender="درخواست تأیید امنیتی",
            text=prompt_text,
            is_user=False,
            is_confirmation=True,
            on_confirm=on_confirm,
            on_cancel=on_cancel
        )
        self._append_bubble(bubble)

    def _append_bubble(self, bubble: MessageBubble):
        # Insert before bottom stretch
        idx = max(0, self.container_layout.count() - 1)
        self.container_layout.insertWidget(idx, bubble)
        # Scroll to bottom
        self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        )
