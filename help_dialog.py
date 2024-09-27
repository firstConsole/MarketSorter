import os
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTextEdit
from PyQt6.QtGui import QIcon


class HelpDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Помощь")
        self.setFixedSize(640, 480)
        self.setWindowIcon(QIcon("icon.png"))

        layout = QVBoxLayout()

        help_file_path = "help.txt"
        if os.path.exists(help_file_path):
            with open(help_file_path, "r", encoding="utf-8") as file:
                help_text = file.read()
        else:
            help_text = "Файл справки не найден."

        text_edit = QTextEdit()
        text_edit.setText(help_text)
        text_edit.setReadOnly(True)  # Сделать текст только для чтения

        layout.addWidget(text_edit)
        self.setLayout(layout)