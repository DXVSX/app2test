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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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

    def get_meters_per_pixel(self, zoom_level):
        # Средний радиус Земли (метры)
        earth_radius = 6378137
        return (2 * 3.14159 * earth_radius) / (256 * 2 ** zoom_level)

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
        options.add_argument("--start-maximized")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")

        try:
            self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=options)
            self.driver.get("https://www.google.com/maps")

            self.hide_elements()

            time.sleep(2)  # Подождите, пока страница загрузится
        except Exception as e:
            self.log(f"Error initializing WebDriver: {e}")
            self.show_popup("Error", f"Error initializing WebDriver: {e}")

    def hide_elements(self):
        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.presence_of_element_located((By.ID, 'scene')))

        elements_to_hide = [
            '//*[@id="omnibox-container"]',
            '//*[@id="assistive-chips"]/div',
            '//*[@id="gb"]',
            '//*[@id="minimap"]/div/div[2]',
            '//*[@id="runway-expand-button"]',
            '//*[@id="content-container"]/div[26]',
            '//*[@id="content-container"]/div[23]/div[1]/div[3]',
            '//*[@id="QA0Szd"]/div/div',
        ]

        for xpath in elements_to_hide:
            try:
                element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                self.driver.execute_script("arguments[0].style.display = 'none';", element)
            except Exception as e:
                self.log(f"Error hiding element {xpath}: {e}")

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
            self.hide_elements()  # Скрываем элементы после обновления URL
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
            self.hide_elements()  # Скрываем элементы после загрузки новой карты
            time.sleep(2)  # Ожидание обновления карты

            # Задаем название файла с текущим временем
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(save_path, f"screenshot_{timestamp}.png")

            # Получаем текущий размер экрана
            screen_width, screen_height = pyautogui.size()

            # Центральные координаты для области скриншота
            center_x = screen_width // 2
            center_y = screen_height // 1.75

            # Задаем границы для скриншота 1919x841 с центрированием
            left = center_x - (1919 // 2)
            top = center_y - (841 // 2)
            right = left + 1919
            bottom = top + 841

            # Делаем скриншот области
            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            screenshot.save(filename)
            self.log(f"Screenshot saved: {filename}")
        except Exception as e:
            self.log(f"Error taking screenshot: {e}")
            self.show_popup("Error", f"Error taking screenshot: {e}")

    def capture_all_houses(self, rect, save_path):
        try:
            # Преобразуем координаты из строк в числа
            top_left_lat, top_left_lon = self.parse_coords(self.top_left_coords)
            bottom_left_lat, bottom_left_lon = self.parse_coords(self.bottom_left_coords)
            top_right_lat, top_right_lon = self.parse_coords(self.top_right_coords)

            # Получаем размер экрана для расчета шага перемещения
            screen_width, screen_height = pyautogui.size()

            # Рассчитываем метры на один пиксель на текущем уровне зума
            meters_per_pixel = self.get_meters_per_pixel(self.current_zoom)

            # Рассчитываем шаги перемещения в широте и долготе (ширина и высота снимка в метрах)
            step_lon = meters_per_pixel * screen_width / 111320  # 1 градус долготой ~ 111.32 км
            step_lat = meters_per_pixel * screen_height / 110540  # 1 градус широтой ~ 110.54 км

            # Уменьшаем шаг по широте для улучшения покрытия
            step_lat *= 0.8  # Уменьшаем шаг по широте на 20%

            # Рассчитываем количество шагов по широте и долготе
            num_steps_x = int((top_right_lon - top_left_lon) / step_lon) + 1  # +1 для включения границы
            num_steps_y = int((top_left_lat - bottom_left_lat) / step_lat) + 1  # +1 для включения границы

            # Перемещение слева направо и сверху вниз
            current_lat = top_left_lat
            while current_lat > bottom_left_lat:
                current_lon = top_left_lon
                while current_lon <= top_right_lon:
                    self.move_to_and_capture(current_lat, current_lon, save_path)
                    current_lon += step_lon

                # Переход на следующую строку
                current_lat -= step_lat  # Перемещаемся вниз после завершения сканирования по широте

            # Дополнительный захват центра области
            center_lat = (top_left_lat + bottom_left_lat) / 2
            center_lon = (top_left_lon + top_right_lon) / 2
            self.move_to_and_capture(center_lat, center_lon, save_path)

            # Захватываем крайние точки для проверки, что граница была захвачена
            self.move_to_and_capture(top_left_lat, top_left_lon, save_path)  # Верхний левый угол
            self.move_to_and_capture(top_left_lat, top_right_lon, save_path)  # Верхний правый угол
            self.move_to_and_capture(bottom_left_lat, top_left_lon, save_path)  # Нижний левый угол
            self.move_to_and_capture(bottom_left_lat, top_right_lon, save_path)  # Нижний правый угол

        except Exception as e:
            self.log(f"Error during house capture: {e}")
            self.show_popup("Error", f"Error during house capture: {e}")

    def parse_coords(self, coords_str):
        try:
            lat, lon = map(float, coords_str.split(','))
            return lat, lon
        except ValueError:
            self.log(f"Invalid coordinates format: {coords_str}")
            return None, None


class MainApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Screenshot Application")
        self.setGeometry(100, 100, 600, 400)

        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)

        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("Enter code")

        self.village_input = QLineEdit()
        self.village_input.setPlaceholderText("Enter village name")

        self.zoom_input = QLineEdit()
        self.zoom_input.setPlaceholderText("Enter zoom level")

        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.start_screenshot)

        layout = QVBoxLayout()
        layout.addWidget(self.log_widget)
        layout.addWidget(self.prefix_input)
        layout.addWidget(self.village_input)
        layout.addWidget(self.zoom_input)
        layout.addWidget(self.submit_button)
        self.setLayout(layout)

        self.screenshot_app = ScreenshotApp(
            self.log_widget,
            self.prefix_input,
            self.village_input,
            self.zoom_input
        )

    def start_screenshot(self):
        self.screenshot_app.start_screenshot()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_app = MainApp()
    main_app.show()
    sys.exit(app.exec_())
