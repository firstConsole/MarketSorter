import sys
import os
import subprocess
import zipfile
import requests
import format_service as fs
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QHBoxLayout
from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtCore import Qt, QFileInfo, QTimer
from PyQt6.QtGui import QIcon, QPixmap, QDragEnterEvent, QDropEvent, QMovie
from PyQt6 import QtCore
from help_dialog import HelpDialog

version = "1.3"

class FileDropWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setFixedSize(800, 600)
        self.setWindowTitle(f"Markets Sorter v {version}")
        self.setWindowIcon(QIcon("icon.png"))

        self.setAcceptDrops(True)

        layout = QVBoxLayout()

        self.image_title = QLabel(self)
        self.image_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pixmap_top = QPixmap("title_image.png")
        self.image_title.setPixmap(pixmap_top)
        layout.addWidget(self.image_title, alignment=Qt.AlignmentFlag.AlignCenter)

        self.label = QLabel("Переместите файл с товарами...", self)
        self.label.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.button = QPushButton("+", self)
        self.button.setFixedSize(700, 300)
        self.button.setStyleSheet("font-size: 100px;")
        self.button.clicked.connect(self.open_file_dialog)

        file_layout = QHBoxLayout()

        self.delete_button = QPushButton("✖", self)
        self.delete_button.setFixedSize(30, 30)
        self.delete_button.setStyleSheet("font-size: 20px; color: red; background-color: transparent; border: none;")
        self.delete_button.clicked.connect(self.clear_file)
        self.delete_button.hide()

        file_layout.addWidget(self.label, alignment=Qt.AlignmentFlag.AlignCenter)
        file_layout.addWidget(self.delete_button, alignment=Qt.AlignmentFlag.AlignLeading)

        self.process_button = QPushButton("Начать обработку", self)
        self.process_button.setFixedSize(170, 50)
        self.process_button.setStyleSheet("font-size: 16px;")
        self.process_button.clicked.connect(self.start_processing)
        self.process_button.setEnabled(False)

        layout.addWidget(self.button, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addLayout(file_layout)
        layout.addWidget(self.process_button, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()

        self.help_button = QPushButton("?", self)
        self.help_button.setFixedSize(50, 50)
        self.help_button.setStyleSheet("font-size: 20px;")
        self.help_button.clicked.connect(self.show_help)

        self.update_button = QPushButton(self)
        self.update_button.setIcon(QIcon("update_icon.png"))
        self.update_button.setIconSize(QtCore.QSize(40, 40))
        self.update_button.setFixedSize(50, 50)
        self.update_button.clicked.connect(self.check_for_updates_manual)

        help_layout = QHBoxLayout()
        help_layout.addStretch()
        help_layout.addWidget(self.update_button, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)
        help_layout.addWidget(self.help_button, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)

        layout.addLayout(help_layout)

        self.setLayout(layout)

        self.file_path = None
        self.save_path = None

        self.loading_label = QLabel(self)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_movie = QMovie("loading.gif")
        self.loading_label.setMovie(self.loading_movie)
        self.loading_label.setFixedSize(50, 50)
        self.loading_label.hide()

        layout.addWidget(self.loading_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.finish_processing)

    def check_for_updates_manual(self):
        try:
            response = requests.get("https://static.insales-cdn.com/files/1/1095/36529223/original/update.json")
            latest_version_info = response.json()
            print(latest_version_info)
            latest_version = latest_version_info['version']
            download_url = latest_version_info['download_url']

            if latest_version > version:
                self.prompt_update(download_url)
            else:
                QMessageBox.information(
                    self,
                    "Проверка обновлений",
                    "Последняя версия! Новее я пока не придумал...",
                    QMessageBox.StandardButton.Ok
                )
        except Exception as e:
            print(f"Ошибка проверки обновлений: {e}")
            QMessageBox.critical(
                self,
                "Ошибка обновления",
                f"Ошибка проверки обновлений: {e}",
                QMessageBox.StandardButton.Ok
            )

    def prompt_update(self, download_url):
        reply = QMessageBox.question(
            self,
            'Обновление доступно',
            "Доступна новая версия приложения. Хотите обновить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.download_update(download_url)

    def download_update(self, download_url):
        try:
            response = requests.get(download_url)
            with open('update.zip', 'wb') as file:
                file.write(response.content)
            print("Обновление загружено.")
            self.apply_update()
        except Exception as e:
            print(f"Ошибка загрузки обновления: {e}")

    def apply_update(self):
        try:
            with zipfile.ZipFile('update.zip', 'r') as zip_ref:
                zip_ref.extractall('.')
            installer_path = './setup.exe'

            if os.path.exists(installer_path):
                subprocess.run([installer_path], check=True)

                os.remove('update.zip')
                print("Обновление успешно применено и архив удален.")
            else:
                print("Установочный файл не найден.")
                QMessageBox.critical(
                    self,
                    "Ошибка",
                    "Не найден установочный файл после распаковки.",
                    QMessageBox.StandardButton.Ok
                )

        except Exception as e:
            print(f"Ошибка применения обновления: {e}")
            QMessageBox.critical(
                self,
                "Ошибка обновления",
                f"Произошла ошибка при применении обновления: {e}",
                QMessageBox.StandardButton.Ok
            )

    def open_file_dialog(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Выберите файл", "",
                                                   "Файлы Excel (*.xls *.xlsx);;CSV Files (*.csv)")
        if file_path:
            self.set_file(file_path)

    def set_file(self, file_path):
        self.file_path = file_path
        file_info = QFileInfo(file_path)
        file_name = file_info.fileName()
        self.label.setText(f"Файл загружен: {file_name}")
        self.delete_button.show()
        self.process_button.setEnabled(True)

    def clear_file(self):
        self.file_path = None
        self.label.setText("Переместите файл с товарами в формате xls, xlsx или csv...")
        self.delete_button.hide()
        self.process_button.setEnabled(False)

    def start_processing(self):
        try:
            if self.file_path:
                self.loading_label.show()
                self.loading_movie.start()

                print(f"Начинаем обработку файла: {self.file_path}")

                result = fs.file_format(self.file_path)

                if isinstance(result, tuple):
                    if len(result) == 2:
                        article_sorted_file, stickers_sorted_file = result

                        save_path_article = self.save_file_dialog(article_sorted_file)
                        if save_path_article:
                            os.rename(article_sorted_file, save_path_article)
                            print(f"Файл с сортировкой по артикулу сохранен как: {save_path_article}")
                            self.label.setText(f"Файл с сортировкой по артикулу сохранен как: {save_path_article}")
                        else:
                            self.label.setText("Сохранение файла с сортировкой по артикулу отменено.")

                        save_path_stickers = self.save_file_dialog(stickers_sorted_file)
                        if save_path_stickers:
                            os.rename(stickers_sorted_file, save_path_stickers)
                            print(f"Файл с сортировкой по стикеру сохранен как: {save_path_stickers}")
                            self.label.setText(f"Файл с сортировкой по стикеру сохранен как: {save_path_stickers}")
                        else:
                            self.label.setText("Сохранение файла с сортировкой по стикеру отменено.")

                elif isinstance(result, str):
                    final_file_path = result
                    save_path = self.save_file_dialog(final_file_path)
                    if save_path:
                        os.rename(final_file_path, save_path)
                        print(f"Файл сохранен как: {save_path}")
                        self.label.setText(f"Файл сохранен как: {save_path}")
                    else:
                        self.label.setText("Сохранение файла отменено.")
                else:
                    self.label.setText("Ошибка: Получен неверный формат файла.")
                    return
            else:
                print("Путь к файлу отсутствует.")
        except Exception as e:
            print(f"Ошибка при обработке файла: {e}")
            self.label.setText(f"Ошибка при обработке файла: {e}")
        finally:
            self.finish_processing()

    def finish_processing(self):
        self.loading_movie.stop()
        self.loading_label.hide()
        self.process_button.setEnabled(False)
        self.delete_button.hide()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if all(url.toLocalFile().endswith(('.xls', '.xlsx', '.csv')) for url in urls):
                event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            self.set_file(file_path)

    def save_file_dialog(self, file):
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить файл как",
            file,
            "Excel Files (*.xls *.xlsx);;CSV Files (*.csv)"
        )
        return save_path

    def show_help(self):
        help_dialog = HelpDialog()
        help_dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileDropWindow()
    window.show()
    sys.exit(app.exec())
