"""
Chat Widget and Message Bubbles for Persian Excel Assistant.
Features deep cosmic glassmorphic bubbles and custom background wallpaper.
"""
import html
import os
from pathlib import Path
import re
import time
from typing import Callable, Optional

from PySide6.QtCore import Qt, Signal, QEvent
from PySide6.QtGui import QPainter, QPixmap, QColor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QScrollArea
)

# Resolve background wallpaper path
RESOURCES_DIR = Path(__file__).parent / "resources"
BACKGROUND_IMAGE_PATH = (RESOURCES_DIR / "background.jpg").as_posix()


def format_markdown_to_rich_text(text: str) -> str:
    """Converts basic markdown (bold, code, line breaks) to clean Qt HTML."""
    if not text:
        return ""
    # Escape HTML special chars first
    escaped = html.escape(text)

    # Convert **bold** to <b>bold</b>
    escaped = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', escaped)

    # Convert `code` to formatted span
    escaped = re.sub(
        r'`(.+?)`',
        r'<span style="background-color: rgba(30, 58, 138, 0.4); color: #38BDF8; padding: 2px 6px; border-radius: 4px; font-family: Consolas, monospace; border: 1px solid #1D4ED8;">\1</span>',
        escaped
    )

    # Convert newlines to <br>
    escaped = escaped.replace('\n', '<br>')
    return f"<div style='line-height: 1.6; font-size: 13px; color: #FFFFFF;'>{escaped}</div>"


class MessageBubble(QFrame):
    """Single chat bubble widget with RTL text and custom glassmorphism styling."""

    def __init__(self, sender: str, text: str, is_user: bool = False,
                 is_confirmation: bool = False, on_confirm: Optional[Callable] = None,
                 on_cancel: Optional[Callable] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("messageBubble")
        self.is_user = is_user
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        # Header: sender name + time
        time_str = time.strftime("%H:%M")
        header_text = f"👤 {sender} ({time_str})" if is_user else f"🤖 {sender} ({time_str})"
        header_lbl = QLabel(header_text)
        header_color = "#38BDF8" if is_user else "#60A5FA"
        header_lbl.setStyleSheet(f"font-size: 11px; color: {header_color}; font-weight: bold; border: none; background: transparent;")
        layout.addWidget(header_lbl)

        # Body text (Rich Text)
        body_lbl = QLabel()
        body_lbl.setTextFormat(Qt.RichText)
        body_lbl.setText(format_markdown_to_rich_text(text))
        body_lbl.setWordWrap(True)
        body_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        body_lbl.setStyleSheet("border: none; background: transparent; color: #FFFFFF;")
        layout.addWidget(body_lbl)

        # Set bubble styles using specific ID selector to avoid child QLabel border inheritance
        if is_confirmation:
            self.setStyleSheet("""
                QFrame#messageBubble {
                    background-color: rgba(45, 22, 10, 0.92);
                    border: 1.5px solid #F59E0B;
                    border-radius: 14px;
                }
                QLabel {
                    border: none;
                    background: transparent;
                }
            """)
            btn_box = QHBoxLayout()
            btn_box.setSpacing(10)
            btn_confirm = QPushButton("✅ تأیید و اجرا")
            btn_confirm.setStyleSheet(
                "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981); "
                "color: white; padding: 7px 16px; border-radius: 8px; font-weight: bold; border: 1px solid #10B981;"
            )
            if on_confirm:
                btn_confirm.clicked.connect(on_confirm)

            btn_cancel = QPushButton("❌ انصراف")
            btn_cancel.setStyleSheet(
                "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #991B1B, stop:1 #DC2626); "
                "color: white; padding: 7px 16px; border-radius: 8px; font-weight: bold; border: 1px solid #EF4444;"
            )
            if on_cancel:
                btn_cancel.clicked.connect(on_cancel)

            btn_box.addWidget(btn_confirm)
            btn_box.addWidget(btn_cancel)
            btn_box.addStretch()
            layout.addLayout(btn_box)

        elif is_user:
            self.setStyleSheet("""
                QFrame#messageBubble {
                    background-color: rgba(12, 25, 58, 0.92);
                    border: 1.5px solid #2563EB;
                    border-radius: 14px;
                }
                QLabel {
                    border: none;
                    background: transparent;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame#messageBubble {
                    background-color: rgba(8, 16, 38, 0.88);
                    border: 1.5px solid #1D4ED8;
                    border-radius: 14px;
                }
                QLabel {
                    border: none;
                    background: transparent;
                }
            """)


class ChatScrollArea(QScrollArea):
    """Custom scroll area that paints the cosmic wave and HM logo backdrop."""

    def __init__(self, bg_image_path: Optional[str] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("chatScrollArea")
        self._bg_pixmap: Optional[QPixmap] = None

        if bg_image_path and os.path.exists(bg_image_path):
            try:
                self._bg_pixmap = QPixmap(bg_image_path)
            except Exception:
                self._bg_pixmap = None

        self.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj is self.viewport() and event.type() == QEvent.Paint:
            try:
                painter = QPainter(self.viewport())
                painter.fillRect(self.viewport().rect(), QColor("#020617"))
                if self._bg_pixmap and not self._bg_pixmap.isNull():
                    vp_w = self.viewport().width()
                    vp_h = self.viewport().height()
                    scaled = self._bg_pixmap.scaled(
                        vp_w, vp_h,
                        Qt.KeepAspectRatioByExpanding,
                        Qt.SmoothTransformation
                    )
                    x = (vp_w - scaled.width()) // 2
                    y = max(0, vp_h - scaled.height())
                    painter.drawPixmap(x, y, scaled)
                painter.end()
            except Exception:
                pass
        return super().eventFilter(obj, event)


class ChatWidget(QWidget):
    """Main scrollable chat feed with cosmic background backdrop."""

    confirm_action_signal = Signal(object)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setLayoutDirection(Qt.RightToLeft)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Scroll Area with custom wallpaper
        self.scroll_area = ChatScrollArea(BACKGROUND_IMAGE_PATH, self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setStyleSheet(
            f"QScrollArea#chatScrollArea {{"
            f"  background-color: #020617;"
            f"  background-image: url('{BACKGROUND_IMAGE_PATH}');"
            f"  background-position: center bottom;"
            f"  background-repeat: no-repeat;"
            f"  border: none;"
            f"}}"
            f"QScrollArea#chatScrollArea > QWidget > QWidget {{"
            f"  background: transparent;"
            f"}}"
        )

        self.container = QWidget()
        self.container.setStyleSheet("background: transparent;")
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(16, 16, 16, 16)
        self.container_layout.setSpacing(12)
        self.container_layout.addStretch()

        self.scroll_area.setWidget(self.container)
        main_layout.addWidget(self.scroll_area)

    def add_user_message(self, text: str):
        bubble = MessageBubble(sender="MIYUM", text=text, is_user=True)
        self._append_bubble(bubble)

    def add_assistant_message(self, text: str):
        bubble = MessageBubble(sender="دستیار هوشمند Microsoft Excel", text=text, is_user=False)
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
