import sys
import os
import time
import pyautogui
import subprocess
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QMainWindow, QRubberBand, QTextEdit, \
    QVBoxLayout, QWidget, QPushButton, QLineEdit, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QRect, QPoint, QSize
from PIL import ImageGrab
from datetime import datetime
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


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
            self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=options)
        except Exception as e:
            self.log(f"Error initializing WebDriver: {e}")
            self.show_popup("Error", f"Error initializing WebDriver: {e}")

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
        try:
            self.take_screenshot()
        except Exception as e:
            self.log(f"Error taking screenshot: {e}")
            self.show_popup("Error", f"Error taking screenshot: {e}")

    def take_screenshot(self):
        rect = QRect(self.start_point, self.end_point).normalized()
        self.hide()

        # Основная папка
        base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'MapsMarkup')

        # Папка с кодом и названием села
        village_name = self.village_input.text().strip()
        save_path = os.path.join(base_folder, self.prefix_input.text().strip(), village_name)
        os.makedirs(save_path, exist_ok=True)

        # Папка для CSV и Excel файлов
        csv_folder = os.path.join(save_path, 'CSV')
        excel_folder = os.path.join(save_path, 'Excel')
        os.makedirs(csv_folder, exist_ok=True)
        os.makedirs(excel_folder, exist_ok=True)

        try:
            self.perform_corner_actions(rect)
            self.capture_all_houses(rect, save_path)
        except Exception as e:
            self.log(f"Error during screenshot processing: {e}")
            self.show_popup("Error", f"Error during screenshot processing: {e}")
        finally:
            self.rubber_band.hide()

    def perform_corner_actions(self, rect):
        pyautogui.sleep(1)

        try:
            # Нажимаем правую кнопку мыши в верхнем левом углу
            pyautogui.moveTo(rect.left() + 10, rect.top() + 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.top_left_coords = self.get_coords_from_maps()
            self.log(f"Top-left coordinates: {self.top_left_coords}")

            # Верхний правый угол
            pyautogui.moveTo(rect.right() - 10, rect.top() + 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.top_right_coords = self.get_coords_from_maps()
            self.log(f"Top-right coordinates: {self.top_right_coords}")

            # Нажимаем правую кнопку мыши в нижнем правом углу
            pyautogui.moveTo(rect.right() - 10, rect.bottom() - 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.bottom_right_coords = self.get_coords_from_maps()
            self.log(f"Bottom-right coordinates: {self.bottom_right_coords}")

            # Нижний левый угол
            pyautogui.moveTo(rect.left() + 10, rect.bottom() - 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.bottom_left_coords = self.get_coords_from_maps()
            self.log(f"Bottom-left coordinates: {self.bottom_left_coords}")

        except Exception as e:
            self.log(f"Error during corner actions: {e}")
            self.show_popup("Error", f"Error during corner actions: {e}")

    def get_coords_from_maps(self):
        try:
            element = self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div')
            coords = element.text
            return coords
        except Exception as e:
            self.log(f"Error: {e}")
            return "Coordinates not found"

    def capture_all_houses(self, rect, save_path):
        """
        Функция делает скриншоты домов, перемещаясь по области и увеличивая масштаб.
        """
        try:
            num_steps = 4  # Количество шагов для перемещения по области
            step_x = rect.width() // num_steps
            step_y = rect.height() // num_steps

            for i in range(num_steps):
                for j in range(num_steps):
                    x_start = rect.left() + i * step_x
                    y_start = rect.top() + j * step_y
                    x_end = x_start + step_x
                    y_end = y_start + step_y

                    # Увеличиваем масштаб
                    self.zoom_in()

                    # Делаем скриншот
                    screenshot = ImageGrab.grab(bbox=(x_start, y_start, x_end, y_end))
                    file_name = datetime.now().strftime(
                        f"screenshot_{self.prefix_input.text().strip()}_{self.village_input.text().strip()}_x_{i}_y_{j}_%Y%m%d_%H%M%S.png")
                    file_path = os.path.join(save_path, file_name)
                    screenshot.save(file_path)

                    # Перемещаемся по странице
                    if j < num_steps - 1:
                        pyautogui.scroll(-100)  # Прокрутка вверх, если нужно

                if i < num_steps - 1:
                    pyautogui.scroll(100 * step_y)  # Прокрутка вниз для следующего ряда
        except Exception as e:
            self.log(f"Error during capturing all houses: {e}")
            self.show_popup("Error", f"Error during capturing all houses: {e}")

    def zoom_in(self):
        """
        Функция увеличивает масштаб карты в Google Maps с помощью JavaScript.
        """
        try:
            zoom_script = "document.querySelector('button[aria-label=\"Zoom in\"]').click();"
            self.driver.execute_script(zoom_script)
            time.sleep(2)  # Даем время браузеру для увеличения
        except Exception as e:
            self.log(f"Error during zooming in: {e}")
            self.show_popup("Error", f"Error during zooming in: {e}")


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
        self.log_widget.append(
            f"Submitted code: {self.input_field.text().strip()}, village: {self.village_field.text().strip()}")
        self.input_field.clear()
        self.village_field.clear()

    def on_tray_icon_click(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.screenshot_app.start_screenshot()


if __name__ == '__main__':
    SystemTrayApp()
