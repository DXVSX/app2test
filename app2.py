from datetime import datetime
import sys
import os
import subprocess
import pyautogui
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QWidget, QVBoxLayout, QTextEdit, \
    QLineEdit, QPushButton, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QRect, QPoint
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from PIL import ImageGrab
import cv2
import numpy as np
import csv
import pyperclip  # Для работы с буфером обмена

prefix = ""
village_name = ""

class ScreenshotApp(QWidget):
    def __init__(self, tray_icon, log_widget, prefix_input, village_input):
        super().__init__()
        self.tray_icon = tray_icon
        self.driver = None
        self.top_left_coords = ""
        self.bottom_right_coords = ""
        self.log_widget = log_widget
        self.prefix_input = prefix_input
        self.village_input = village_input
        self.selection_rect = None
        self.setup_driver()

    def show_popup(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def setup_driver(self):
        options = Options()
        options.add_experimental_option("debuggerAddress", "localhost:9222")
        self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=options)

    def log(self, message):
        self.log_widget.append(message)
        self.log_widget.verticalScrollBar().setValue(self.log_widget.verticalScrollBar().maximum())

    def copy_to_clipboard(self, text):
        pyperclip.copy(text)
        self.log(f"Copied to clipboard: {text}")

    def get_coords_from_maps(self):
        try:
            # Пример кода для получения координат (измените в соответствии с вашим интерфейсом)
            element = self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div')
            coords = element.text
            self.copy_to_clipboard(coords)
            return coords
        except Exception as e:
            self.log(f"Error: {e}")
            return "Coordinates not found"

    def record_coordinates(self):
        try:
            # Копируем координаты топ-левой и нижней правой точек
            self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div').click()
            self.top_left_coords = self.get_coords_from_maps()
            self.log(f"Top-left coordinates: {self.top_left_coords}")

            self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div').click()
            self.bottom_right_coords = self.get_coords_from_maps()
            self.log(f"Bottom-right coordinates: {self.bottom_right_coords}")

            # Сохраняем координаты в CSV файл
            self.save_coordinates()

            # Автоматически захватываем снимок экрана в записанной области
            self.capture_screenshot_in_area()

        except Exception as e:
            self.show_popup("Error", f"An error occurred: {str(e)}")

    def save_coordinates(self):
        base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'MapsMarkup')
        village_name = self.village_input.text().strip()
        save_path = os.path.join(base_folder, prefix, village_name)
        os.makedirs(save_path, exist_ok=True)

        csv_path = os.path.join(save_path, 'coordinates.csv')
        with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Top-Left Coordinates', 'Bottom-Right Coordinates'])
            writer.writerow([self.top_left_coords, self.bottom_right_coords])

        self.show_popup(
            "Coordinates saved",
            f"Coordinates saved in {csv_path}"
        )

    def capture_screenshot_in_area(self):
        try:
            if not self.top_left_coords or not self.bottom_right_coords:
                self.show_popup("Error", "Coordinates are not set.")
                return

            # Преобразуем координаты из строки в числа с плавающей запятой
            lat_top_left, lon_top_left = map(float, self.top_left_coords.split(','))
            lat_bottom_right, lon_bottom_right = map(float, self.bottom_right_coords.split(','))

            # Определяем область для снимка экрана
            x1, y1 = 100, 100  # Временные значения для верхнего левого угла
            x2, y2 = 800, 600  # Временные значения для нижнего правого угла

            # Захватываем снимок экрана указанной области
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))

            # Сохраняем снимок экрана
            base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'MapsMarkup')
            village_name = self.village_input.text().strip()
            save_path = os.path.join(base_folder, prefix, village_name)
            os.makedirs(save_path, exist_ok=True)

            screenshot_filename = f"screenshot_{prefix}_{village_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            screenshot_path = os.path.join(save_path, screenshot_filename)
            screenshot.save(screenshot_path)

            self.show_popup(
                "Screenshot saved",
                f"Screenshot saved in {screenshot_path}"
            )
        except Exception as e:
            self.show_popup("Error", f"An error occurred while taking the screenshot: {str(e)}")


class SystemTrayApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        self.main_window = QWidget()
        self.layout = QVBoxLayout()
        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)
        self.layout.addWidget(self.log_widget)
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Enter prefix (code) here")
        self.layout.addWidget(self.input_field)
        self.village_field = QLineEdit()
        self.village_field.setPlaceholderText("Enter village name here")
        self.layout.addWidget(self.village_field)
        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.on_submit)
        self.layout.addWidget(self.submit_button)
        self.main_window.setLayout(self.layout)
        self.main_window.resize(800, 600)
        self.open_google_maps()
        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self.app)
        self.screenshot_app = ScreenshotApp(self.tray_icon, self.log_widget, self.input_field, self.village_field)
        self.menu = QMenu()
        self.exit_action = QAction("Exit", self.app)
        self.exit_action.triggered.connect(self.app.quit)
        self.menu.addAction(self.exit_action)
        self.tray_icon.setContextMenu(self.menu)
        self.tray_icon.activated.connect(self.on_tray_icon_click)
        self.tray_icon.show()
        self.main_window.show()
        sys.exit(self.app.exec_())

    def open_google_maps(self):
        try:
            command = 'start chrome --remote-debugging-port=9222 "https://www.google.com/maps"'
            subprocess.Popen(command, shell=True)
        except Exception as e:
            print(f"Failed to open Google Maps: {str(e)}")

    def on_submit(self):
        global prefix
        prefix = self.input_field.text().strip()
        village_name = self.village_field.text().strip()
        self.log_widget.append(f"Submitted code: {prefix}, village: {village_name}")
        self.input_field.clear()
        self.village_field.clear()
        self.screenshot_app.record_coordinates()

    def on_tray_icon_click(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.screenshot_app.record_coordinates()


if __name__ == '__main__':
    SystemTrayApp()
