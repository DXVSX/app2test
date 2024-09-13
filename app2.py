import sys
import os
import time
import pyautogui
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QMainWindow, QRubberBand, QTextEdit, \
    QVBoxLayout, QWidget, QPushButton, QLineEdit, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QRect, QPoint, QSize
from PIL import ImageGrab, Image
from PIL.Image import Resampling  # Импортируем Resampling
from datetime import datetime
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import re


class ScreenshotApp(QMainWindow):
    def __init__(self, tray_icon, log_widget, prefix_input, village_input, zoom_input):
        super().__init__()
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.rubber_band = QRubberBand(QRubberBand.Rectangle, self)
        self.tray_icon = tray_icon
        self.driver = None
        self.top_left_coords = ""
        self.top_right_coords = ""
        self.bottom_right_coords = ""
        self.bottom_left_coords = ""
        self.log_widget = log_widget
        self.prefix_input = prefix_input
        self.village_input = village_input
        self.zoom_input = zoom_input
        self.setup_driver()

        # Параметры зума
        self.current_zoom = 17  # Фиксированный уровень зума

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
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")

        try:
            self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=options)
            self.driver.get("https://www.google.com/maps")
            time.sleep(5)  # Подождите, пока страница загрузится
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
            pyautogui.click(button='left')

            # Верхний правый угол
            pyautogui.moveTo(rect.right() - 10, rect.top() + 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.top_right_coords = self.get_coords_from_maps()
            self.log(f"Top-right coordinates: {self.top_right_coords}")
            pyautogui.click(button='left')

            # Нажимаем правую кнопку мыши в нижнем правом углу
            pyautogui.moveTo(rect.right() - 10, rect.bottom() - 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.bottom_right_coords = self.get_coords_from_maps()
            self.log(f"Bottom-right coordinates: {self.bottom_right_coords}")
            pyautogui.click(button='left')

            # Нижний левый угол
            pyautogui.moveTo(rect.left() + 10, rect.bottom() - 10)
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.bottom_left_coords = self.get_coords_from_maps()
            self.log(f"Bottom-left coordinates: {self.bottom_left_coords}")
            pyautogui.click(button='left')

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
        Функция делает скриншоты домов, перемещаясь по области и устанавливая уровень зума на 17.
        """
        try:
            num_steps = 4  # Количество шагов для перемещения по области
            step_x = rect.width() // num_steps
            step_y = rect.height() // num_steps

            # Установка фиксированного уровня зума
            self.set_zoom_level(self.driver, zoom_level=17)

            for i in range(num_steps):
                for j in range(num_steps):
                    x_start = rect.left() + i * step_x
                    y_start = rect.top() + j * step_y
                    x_end = x_start + step_x
                    y_end = y_start + step_y

                    # Перемещаемся и делаем скриншот
                    self.move_to_and_capture(x_start, y_start, save_path)

        except Exception as e:
            self.log(f"Error during capturing all houses: {e}")
            self.show_popup("Error", f"Error during capturing all houses: {e}")

    def move_to_and_capture(self, x, y, save_path):
        """
        Перемещает карту к указанной позиции и делает скриншот, изменяя размер на 1180x700 пикселей.
        """
        try:
            pyautogui.moveTo(x, y)
            pyautogui.click(button='left')
            time.sleep(2)  # Ожидание, пока карта обновится

            # Задаем название файла с текущим временем
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(save_path, f"screenshot_{timestamp}.png")

            # Делаем скриншот
            screenshot = ImageGrab.grab(bbox=(x, y, x + 1180, y + 700))

            # Изменяем размер изображения на 1180x700 пикселей
            screenshot = screenshot.resize((1180, 700), Resampling.LANCZOS)
            screenshot.save(filename)
            self.log(f"Screenshot saved to {filename}")

        except Exception as e:
            self.log(f"Error capturing screenshot at ({x}, {y}): {e}")
            self.show_popup("Error", f"Error capturing screenshot at ({x}, {y}): {e}")

    def set_zoom_level(self, driver, zoom_level):
        latitude, longitude, _ = self.get_coordinates_and_zoom(driver)
        if latitude is not None and longitude is not None:
            new_url = f'https://www.google.com/maps/@{latitude},{longitude},{zoom_level}z'
            driver.get(new_url)
            time.sleep(2)  # Ожидание обновления карты
        else:
            self.log("Не удалось получить текущие координаты. Попробуйте снова.")

    def get_coordinates_and_zoom(self, driver):
        url = driver.current_url
        # Регулярное выражение для извлечения координат и уровня зума
        match = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+),(\d+)z', url)
        if match:
            latitude = float(match.group(1))
            longitude = float(match.group(2))
            zoom_level = int(match.group(3))
            return latitude, longitude, zoom_level
        else:
            return None, None, None


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Screenshot Tool")
        self.setGeometry(100, 100, 400, 300)

        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self)
        tray_menu = QMenu()
        quit_action = QAction("Exit", self)
        quit_action.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # Layout and widgets
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)
        layout.addWidget(self.log_widget)

        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("Enter code (e.g., CM, SL)")
        layout.addWidget(self.prefix_input)

        self.village_input = QLineEdit()
        self.village_input.setPlaceholderText("Enter village name")
        layout.addWidget(self.village_input)

        self.zoom_input = QLineEdit()
        self.zoom_input.setPlaceholderText("Zoom level (fixed to 17)")
        layout.addWidget(self.zoom_input)

        submit_button = QPushButton("Submit")
        submit_button.clicked.connect(self.submit_data)
        layout.addWidget(submit_button)

        self.screenshot_app = ScreenshotApp(self.tray_icon, self.log_widget, self.prefix_input, self.village_input,
                                            self.zoom_input)

    def submit_data(self):
        self.screenshot_app.start_screenshot()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
