import sys
import os
import time
import pyautogui
from PyQt5.QtWidgets import QApplication, QMainWindow, QRubberBand, QTextEdit, QVBoxLayout, QWidget, QPushButton, QLineEdit, QMessageBox
from PyQt5.QtCore import QRect, QPoint, QSize
from PIL import ImageGrab, Image
from PIL.Image import Resampling
from datetime import datetime
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import re


class ScreenshotApp(QMainWindow):
    def __init__(self, log_widget, prefix_input, village_input, zoom_input):
        super().__init__()
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.rubber_band = QRubberBand(QRubberBand.Rectangle, self)
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
            # Верхний левый угол
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

            # Нижний правый угол
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
            # Разбиваем координаты на широту и долготу
            return coords  # Вернем координаты в виде строки
        except Exception as e:
            self.log(f"Error: {e}")
            return ""

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
        match = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+),(\d+)z', url)
        if match:
            latitude = float(match.group(1))
            longitude = float(match.group(2))
            zoom_level = int(match.group(3))
            return latitude, longitude, zoom_level
        else:
            return None, None, None

    def move_to_and_capture(self, latitude, longitude, save_path):
        try:
            # Обновляем URL с новыми координатами
            new_url = f'https://www.google.com/maps/@{latitude},{longitude},{self.current_zoom}z'
            self.driver.get(new_url)
            time.sleep(2)  # Ожидание обновления карты

            # Задаем название файла с текущим временем
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(save_path, f"screenshot_{timestamp}.png")

            # Делаем скриншот
            screenshot = ImageGrab.grab(bbox=(100, 100, 1280, 800))  # Параметры области экрана для скриншота
            screenshot = screenshot.resize((1180, 700), Resampling.LANCZOS)
            screenshot.save(filename)
            self.log(f"Screenshot saved to {filename}")

            # Добавляем паузу перед следующим шагом
            time.sleep(4)  # Время ожидания для загрузки карты с новыми координатами

        except Exception as e:
            self.log(f"Error capturing screenshot at ({latitude}, {longitude}): {e}")
            self.show_popup("Error", f"Error capturing screenshot at ({latitude}, {longitude}): {e}")

    def capture_all_houses(self, rect, save_path, step_lat=0.0020, step_lon=0.0060):
        try:
            # Преобразуем координаты углов в числовой формат
            top_left_lat, top_left_lon = self.parse_coords(self.top_left_coords)
            top_right_lat, top_right_lon = self.parse_coords(self.top_right_coords)
            bottom_right_lat, bottom_right_lon = self.parse_coords(self.bottom_right_coords)
            bottom_left_lat, bottom_left_lon = self.parse_coords(self.bottom_left_coords)

            if None in (top_left_lat, top_left_lon, top_right_lat, top_right_lon, bottom_right_lat, bottom_right_lon, bottom_left_lat, bottom_left_lon):
                raise ValueError("Не удалось получить все координаты углов.")

            num_steps_x = int((top_right_lon - top_left_lon) // step_lon)
            num_steps_y = int((top_left_lat - bottom_left_lat) // step_lat)

            for j in range(num_steps_y):
                # Перемещение слева направо
                for i in range(num_steps_x):
                    current_lat = top_left_lat - j * step_lat
                    current_lon = top_left_lon + i * step_lon
                    self.move_to_and_capture(current_lat, current_lon, save_path)

        except Exception as e:
            self.log(f"Error during house capture: {e}")
            self.show_popup("Error", f"Error during house capture: {e}")

    def parse_coords(self, coords):
        try:
            if coords:
                # Если координаты в формате строки, преобразуем их
                if isinstance(coords, str):
                    return tuple(map(float, coords.split(',')))
                # Если координаты уже в виде кортежа, возвращаем их как есть
                elif isinstance(coords, tuple) and len(coords) == 2:
                    return coords
            return None, None
        except Exception as e:
            self.log(f"Error parsing coordinates: {e}")
            return None, None


def main():
    app = QApplication(sys.argv)
    window = QWidget()
    layout = QVBoxLayout()

    log_widget = QTextEdit()
    layout.addWidget(log_widget)

    prefix_input = QLineEdit()
    layout.addWidget(prefix_input)
    prefix_input.setPlaceholderText("Введите код")

    village_input = QLineEdit()
    layout.addWidget(village_input)
    village_input.setPlaceholderText("Введите название села")

    zoom_input = QLineEdit()
    layout.addWidget(zoom_input)
    zoom_input.setPlaceholderText("Введите лимит зума")

    screenshot_button = QPushButton("Сделать снимок")
    layout.addWidget(screenshot_button)

    window.setLayout(layout)
    window.show()

    app_instance = ScreenshotApp(log_widget, prefix_input, village_input, zoom_input)
    screenshot_button.clicked.connect(app_instance.start_screenshot)

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
