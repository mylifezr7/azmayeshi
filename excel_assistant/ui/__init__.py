"""
User Interface module for Excel AI Assistant.
Provides both PySide6 modern desktop GUI and responsive Web UI with live Excel cell grid preview.
"""
from excel_assistant.ui.controller import AssistantUIController, ChatMessage
from excel_assistant.ui.web_server import start_web_server

__all__ = [
    "AssistantUIController",
    "ChatMessage",
    "start_web_server",
]
