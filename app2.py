import time
import sys
import os
import pyautogui
import subprocess
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QMainWindow, QRubberBand, QTextEdit, \
    QVBoxLayout, QWidget, QPushButton, QLineEdit, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QRect, QPoint, QSize
from PIL import ImageGrab
from datetime import datetime
import cv2
import numpy as np
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from PyQt5.QtWidgets import QMessageBox

prefix = ""
village_name = ""

class ScreenshotApp(QMainWindow):
    def __init__(self, tray_icon, log_widget, prefix_input, village_input):
        super().__init__()
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.rubber_band = QRubberBand(QRubberBand.Rectangle, self)
        self.tray_icon = tray_icon
        self.driver = None
        self.top_left_coords = ""
        self.bottom_right_coords = ""
        self.log_widget = log_widget
        self.prefix_input = prefix_input
        self.village_input = village_input
        self.setup_driver()

    def show_popup(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def start_screenshot(self):
        self.showFullScreen()
        self.setWindowOpacity(0.3)
        self.show()

    def setup_driver(self):
        options = Options()
        options.add_experimental_option("debuggerAddress", "localhost:9222")
        try:
            self.driver = webdriver.Chrome(service=Service('C:/Users/Dan/Desktop/chromedriver.exe'), options=options)
        except Exception as e:
            self.log(f"Failed to setup WebDriver: {str(e)}")
            self.show_popup("Error", f"Failed to setup WebDriver: {str(e)}")

    def log(self, message):
        self.log_widget.append(message)
        self.log_widget.verticalScrollBar().setValue(self.log_widget.verticalScrollBar().maximum())

    def mousePressEvent(self, event):
        self.start_point = event.pos()
        self.rubber_band.setGeometry(QRect(self.start_point, QSize()))
        self.rubber_band.show()

    def mouseMoveEvent(self, event):
        self.rubber_band.setGeometry(QRect(self.start_point, event.pos()).normalized())

    def mouseReleaseEvent(self, event):
        self.end_point = event.pos()
        self.take_screenshot()

    def take_screenshot(self):
        rect = QRect(self.start_point, self.end_point).normalized()
        self.hide()
        screenshot = ImageGrab.grab(bbox=(rect.left(), rect.top(), rect.right(), rect.bottom()))

        # Основная папка
        base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'MapsMarkup')
        
        # Папка с кодом и названием села
        village_name = self.village_input.text().strip()
        save_path = os.path.join(base_folder, prefix, village_name)
        os.makedirs(save_path, exist_ok=True)

        # Папка для CSV и Excel файлов
        csv_folder = os.path.join(save_path, 'CSV')
        excel_folder = os.path.join(save_path, 'Excel')
        os.makedirs(csv_folder, exist_ok=True)
        os.makedirs(excel_folder, exist_ok=True)

        file_name = datetime.now().strftime(f"screenshot_{prefix}_{village_name}_%Y%m%d_%H%M%S.png")
        file_path = os.path.join(save_path, file_name)

        screenshot.save(file_path)

        self.perform_corner_actions(rect)
        self.process_image(file_path, self.top_left_coords, self.bottom_right_coords, csv_folder, excel_folder)
        self.rubber_band.hide()

    def perform_corner_actions(self, rect):
        pyautogui.sleep(1)
        pyautogui.moveTo(rect.left() + 10, rect.top() + 10)
        pyautogui.sleep(1)
        pyautogui.click(button='right')
        pyautogui.sleep(1)
        self.top_left_coords = self.get_coords_from_maps()
        self.log(f"Top-left coordinates: {self.top_left_coords}")
        pyautogui.moveTo(rect.right() - 10, rect.bottom() - 10)
        pyautogui.sleep(1)
        pyautogui.click(button='right')
        pyautogui.sleep(1)
        self.bottom_right_coords = self.get_coords_from_maps()
        self.log(f"Bottom-right coordinates: {self.bottom_right_coords}")

    def get_coords_from_maps(self):
        try:
            element = self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div')
            coords = element.text
            return coords
        except Exception as e:
            self.log(f"Error: {e}")
            return "Coordinates not found"

    def process_image(self, filepath, top_left_coords, bottom_right_coords, csv_folder, excel_folder):
        try:
            lat_top_left, lon_top_left = map(float, top_left_coords.split(','))
            lat_bottom_right, lon_bottom_right = map(float, bottom_right_coords.split(','))

            image = cv2.imread(filepath)
            height, width, _ = image.shape

            target_color = np.array([237, 233, 232])
            mask = cv2.inRange(image, target_color, target_color)
            contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            lat_scale = (lat_top_left - lat_bottom_right) / height
            lon_scale = (lon_bottom_right - lon_top_left) / width

            # Сохранение CSV файла
            csv_path = self.get_unique_filename(csv_folder, 'houses.csv')
            with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(['Marker', 'Lat', 'Long'])
                global prefix
                for i, contour in enumerate(contours, start=1):
                    if cv2.contourArea(contour) > 10:
                        M = cv2.moments(contour)
                        if M['m00'] != 0:
                            cx = int(M['m10'] / M['m00'])
                            cy = int(M['m01'] / M['m00'])
                            lat = lat_top_left - cy * lat_scale
                            lon = lon_top_left + cx * lon_scale
                            marker = f'{prefix}{i}'
                            writer.writerow([marker, lat, lon])

            # Сохранение Excel файла
            excel_path = self.get_unique_filename(excel_folder, 'houses.xlsx')
            data = pd.read_csv(csv_path)
            data.to_excel(excel_path, index=False)

            self.show_popup(
                "Files saved",
                f"CSV and Excel files saved in {csv_folder} and {excel_folder}"
            )
        except Exception as e:
            self.show_popup("Error", f"An error occurred: {str(e)}")

    def get_unique_filename(self, folder, filename):
        base_name, ext = os.path.splitext(filename)
        counter = 1
        unique_filename = os.path.join(folder, filename)
        while os.path.exists(unique_filename):
            unique_filename = os.path.join(folder, f"{base_name}_{counter}{ext}")
            counter += 1
        return unique_filename


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

    def on_tray_icon_click(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.screenshot_app.start_screenshot()


if __name__ == '__main__':
    SystemTrayApp()
