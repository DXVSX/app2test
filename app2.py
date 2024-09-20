import sys
import os
import time
import pyautogui
import cv2
import numpy as np
import math
from PyQt5.QtWidgets import QApplication, QMainWindow, QRubberBand, QTextEdit, QVBoxLayout, QWidget, QPushButton, \
    QLineEdit, QMessageBox, QDialog, QLabel
from PyQt5.QtCore import QRect, QPoint, QSize, Qt
from PIL import ImageGrab
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


class CompletionDialog(QDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedSize(500, 300)  # Устанавливаем желаемый размер окна

        # Центрируем окно на экране
        self.center_on_screen()

        layout = QVBoxLayout()
        label = QLabel(message)
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 16pt;")
        layout.addWidget(label)

        ok_button = QPushButton("OK")
        ok_button.setFixedSize(100, 40)
        ok_button.clicked.connect(self.accept)
        ok_button.setStyleSheet("font-size: 14pt;")

        layout.addWidget(ok_button, alignment=Qt.AlignCenter)
        self.setLayout(layout)

    def center_on_screen(self):
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)


class ScreenshotApp(QMainWindow):
    def __init__(self, log_widget, prefix_input, village_input):
        super().__init__()
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.rubber_band = QRubberBand(QRubberBand.Rectangle, self)
        self.driver = None
        self.top_left_coords = ""
        self.bottom_right_coords = ""
        self.log_widget = log_widget
        self.prefix_input = prefix_input
        self.village_input = village_input
        self.setup_driver()
        self.current_zoom = 19  # Уровень масштаба для лучшей точности
        self.house_coordinates = []

        # Целевой цвет домов в HEX (#E8E9ED)
        target_color_hex = '#E8E9ED'
        # Преобразуем HEX в RGB
        target_color_rgb = tuple(int(target_color_hex[i:i + 2], 16) for i in (1, 3, 5))
        # Преобразуем в BGR для OpenCV
        self.target_color = np.array(target_color_rgb[::-1], dtype=np.uint8)  # Это будет [237, 233, 232]

    def get_meters_per_pixel(self, latitude, zoom_level):
        # Эта функция вычисляет, сколько метров представляет один пиксель при заданной широте и уровне масштаба
        return (156543.03392 * math.cos(math.radians(latitude))) / (2 ** zoom_level)

    def show_popup(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def show_completion_dialog(self, title, message):
        dialog = CompletionDialog(title, message, self)
        dialog.exec_()

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

            time.sleep(2)  # Ожидание загрузки страницы
        except Exception as e:
            self.log(f"Error initializing WebDriver: {e}")
            self.show_popup("Ошибка", f"Ошибка инициализации WebDriver: {e}")

    def hide_elements(self):
        try:
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

        except Exception as e:
            self.log(f"Error during hide_elements: {e}")

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
            self.show_popup("Ошибка", f"Ошибка при создании скриншота: {e}")

    def take_screenshot(self):
        rect = QRect(self.start_point, self.end_point).normalized()
        self.hide()

        base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'MapsMarkup')

        village_name = self.village_input.text().strip()
        save_path = os.path.join(base_folder, self.prefix_input.text().strip(), village_name)
        os.makedirs(save_path, exist_ok=True)

        csv_folder = os.path.join(save_path, 'CSV')
        excel_folder = os.path.join(save_path, 'Excel')
        os.makedirs(csv_folder, exist_ok=True)
        os.makedirs(excel_folder, exist_ok=True)

        try:
            self.perform_corner_actions(rect)
            self.capture_all_houses(save_path)
            self.mark_houses(csv_folder, excel_folder)
        except Exception as e:
            self.log(f"Error during screenshot processing: {e}")
            self.show_popup("Ошибка", f"Ошибка при обработке скриншота: {e}")
        finally:
            self.rubber_band.hide()

    def perform_corner_actions(self, rect):
        pyautogui.sleep(1)

        try:
            pyautogui.moveTo(rect.left(), rect.top())
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.top_left_coords = self.get_coords_from_maps()
            self.log(f"Top-left coordinates: {self.top_left_coords}")

            pyautogui.moveTo(rect.right(), rect.bottom())
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.bottom_right_coords = self.get_coords_from_maps()
            self.log(f"Bottom-right coordinates: {self.bottom_right_coords}")

        except Exception as e:
            self.log(f"Error during corner actions: {e}")
            self.show_popup("Ошибка", f"Ошибка при получении координат: {e}")

    def get_coords_from_maps(self):
        try:
            element = self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div')
            coords = element.text
            return coords
        except Exception as e:
            self.log(f"Error: {e}")
            return ""

    def is_duplicate(self, new_lat, new_lon, house_coordinates):
        # Устанавливаем допустимое расстояние между домами (в метрах)
        allowed_distance = 10  # 10 метров

        # Переводим координаты в радианы
        new_lat_rad = math.radians(new_lat)
        new_lon_rad = math.radians(new_lon)

        for lat, lon in house_coordinates:
            lat_rad = math.radians(lat)
            lon_rad = math.radians(lon)
            # Используем формулу haversine для вычисления расстояния между двумя точками на сфере
            dlat = lat_rad - new_lat_rad
            dlon = lon_rad - new_lon_rad
            a = math.sin(dlat / 2) ** 2 + math.cos(new_lat_rad) * math.cos(lat_rad) * math.sin(dlon / 2) ** 2
            c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
            distance = 6371000 * c  # Радиус Земли ~6371 км

            if distance < allowed_distance:
                return True
        return False

    def move_to_and_capture(self, latitude, longitude, save_path):
        try:
            meters_per_pixel = self.get_meters_per_pixel(latitude, self.current_zoom)

            new_url = f'https://www.google.com/maps/@{latitude},{longitude},{self.current_zoom}z'
            self.driver.get(new_url)
            self.hide_elements()
            time.sleep(2)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(save_path, f"screenshot_{timestamp}.png")

            self.driver.save_screenshot(filename)
            self.log(f"Screenshot saved: {filename}")

            house_coords_on_image = self.detect_houses_on_map(
                filename, latitude, longitude, meters_per_pixel
            )

            for house_lat, house_lon in house_coords_on_image:
                if not self.is_duplicate(house_lat, house_lon, self.house_coordinates):
                    self.house_coordinates.append((house_lat, house_lon))

        except Exception as e:
            self.log(f"Error taking screenshot: {e}")
            self.show_popup("Ошибка", f"Ошибка при создании скриншота: {e}")

    def detect_houses_on_map(self, image_path, latitude, longitude, meters_per_pixel):
        image = cv2.imread(image_path)
        if image is None:
            self.log(f"Error loading image {image_path}")
            return []

        # Get image dimensions
        height, width, _ = image.shape

        # Target color in BGR
        target_color_bgr = self.target_color  # Это numpy массив со значениями [237, 233, 232]

        # Tolerance in color values (set tolerance to 5 to highlight similar colors)
        tolerance = 5
        lower_bound = np.clip(target_color_bgr - tolerance, 0, 255)
        upper_bound = np.clip(target_color_bgr + tolerance, 0, 255)

        # Create mask with tolerance
        mask = cv2.inRange(image, lower_bound, upper_bound)

        # Save mask for debugging
        mask_filename = image_path.replace('.png', '_mask.png')
        cv2.imwrite(mask_filename, mask)
        self.log(f"Mask saved: {mask_filename}")

        # Perform morphological closing to bridge small gaps or lines between houses
        kernel_size = 5  # Увеличили размер ядра для объединения домов, разделенных до 1.5 пикселя
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
        closed_mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)

        # Save the closed mask for debugging
        closed_mask_filename = image_path.replace('.png', '_closed_mask.png')
        cv2.imwrite(closed_mask_filename, closed_mask)
        self.log(f"Closed mask saved: {closed_mask_filename}")

        # Use connectedComponentsWithStats to find connected regions
        num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(closed_mask, connectivity=8)
        self.log(f"Found {num_labels - 1} connected components")

        # Save labeled image for debugging
        label_hue = np.uint8(179 * labels / np.max(labels))
        blank_ch = 255 * np.ones_like(label_hue)
        labeled_img = cv2.merge([label_hue, blank_ch, blank_ch])
        labeled_img = cv2.cvtColor(labeled_img, cv2.COLOR_HSV2BGR)
        labeled_img[label_hue == 0] = 0

        labeled_filename = image_path.replace('.png', '_labeled.png')
        cv2.imwrite(labeled_filename, labeled_img)
        self.log(f"Labeled image saved: {labeled_filename}")

        house_coordinates = []

        for label in range(1, num_labels):
            area = stats[label, cv2.CC_STAT_AREA]
            if 100 < area < 10000:
                cx, cy = centroids[label]
                dx_pixels = cx - (width / 2)
                dy_pixels = cy - (height / 2)

                dx_meters = dx_pixels * meters_per_pixel
                dy_meters = dy_pixels * meters_per_pixel

                delta_lat = -dy_meters / 110540  # 1 градус широты ~110.54 км
                delta_lon = dx_meters / (111320 * math.cos(math.radians(latitude)))

                house_lat = latitude + delta_lat
                house_lon = longitude + delta_lon

                # Добавляем координаты дома
                house_coordinates.append((house_lat, house_lon))

        return house_coordinates

    def capture_all_houses(self, save_path):
        try:
            self.house_coordinates = []
            visited_positions = set()

            top_left_lat, top_left_lon = self.parse_coords(self.top_left_coords)
            bottom_right_lat, bottom_right_lon = self.parse_coords(self.bottom_right_coords)

            screen_width, screen_height = pyautogui.size()

            current_lat = top_left_lat

            while current_lat > bottom_right_lat:
                current_lon = top_left_lon
                meters_per_pixel = self.get_meters_per_pixel(current_lat, self.current_zoom)

                # Вычисляем шаги с учетом текущей широты
                step_lon = (meters_per_pixel * screen_width / (111320 * math.cos(math.radians(current_lat))))
                step_lat = (meters_per_pixel * screen_height / 110540)

                while current_lon < bottom_right_lon:
                    # Округляем координаты до 6 знаков после запятой для проверки на дубликаты позиций
                    position_key = (round(current_lat, 6), round(current_lon, 6))
                    if position_key not in visited_positions:
                        self.move_to_and_capture(current_lat, current_lon, save_path)
                        visited_positions.add(position_key)
                    else:
                        self.log(f"Skipping duplicate position at lat: {current_lat}, lon: {current_lon}")
                    current_lon += step_lon

                current_lat -= step_lat

        except Exception as e:
            self.log(f"Error during house capture: {e}")
            self.show_popup("Ошибка", f"Ошибка при захвате домов: {e}")

    def mark_houses(self, csv_folder, excel_folder):
        try:
            csv_filename = os.path.join(csv_folder, 'houses.csv')
            with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['Marker', 'Latitude', 'Longitude']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                for idx, (lat, lon) in enumerate(self.house_coordinates, start=1):
                    # Формируем маркер в формате CM001, CM002 и т.д.
                    marker = f"{self.prefix_input.text().strip()}{idx:03d}"
                    writer.writerow({
                        'Marker': marker,
                        'Latitude': f"{lat:.14f}",
                        'Longitude': f"{lon:.14f}"
                    })

            df = pd.read_csv(csv_filename)
            excel_filename = os.path.join(excel_folder, 'houses.xlsx')
            df.to_excel(excel_filename, index=False)

            self.log(f"Houses marked and saved to {csv_filename} and {excel_filename}")

            # Добавляем уведомление об окончании процесса
            self.show_completion_dialog("Процесс завершен", "Процесс успешно завершен. Файлы сохранены.")

        except Exception as e:
            self.log(f"Error marking houses: {e}")
            self.show_popup("Ошибка", f"Ошибка при сохранении домов: {e}")

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

        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.start_screenshot)

        layout = QVBoxLayout()
        layout.addWidget(self.log_widget)
        layout.addWidget(self.prefix_input)
        layout.addWidget(self.village_input)
        layout.addWidget(self.submit_button)
        self.setLayout(layout)

        self.screenshot_app = ScreenshotApp(
            self.log_widget,
            self.prefix_input,
            self.village_input
        )

    def start_screenshot(self):
        self.screenshot_app.start_screenshot()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_app = MainApp()
    main_app.show()
    sys.exit(app.exec_())
