import sys
import os
import time
import cv2
import numpy as np
import math
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QRubberBand, QTextEdit, QVBoxLayout,
    QWidget, QPushButton, QLineEdit, QMessageBox, QDialog, QLabel
)
from PyQt5.QtCore import QRect, QPoint, QSize, Qt
from datetime import datetime
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from sklearn.cluster import DBSCAN

class CompletionDialog(QDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedSize(700, 300)  # Set the desired window size
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint | Qt.WindowModal)
        self.center_on_screen()

        layout = QVBoxLayout()
        label = QLabel(message)
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 12pt;")
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
        self.current_zoom = 19  # Zoom level for better accuracy
        self.house_coordinates = []

        # Target house color in HEX (#E8E9ED)
        target_color_hex = '#E8E9ED'
        # Convert HEX to RGB
        target_color_rgb = tuple(int(target_color_hex[i:i + 2], 16) for i in (1, 3, 5))
        # Convert to BGR for OpenCV
        self.target_color = np.array(target_color_rgb[::-1], dtype=np.uint8)

        # Filtering parameters
        self.min_area = 200     # Minimum area of the object
        self.max_area = 50000   # Increased maximum area to include larger houses
        self.min_aspect_ratio = 0.2  # Minimum aspect ratio
        self.max_aspect_ratio = 2.0  # Maximum aspect ratio

        # Thresholds to distinguish between sheds and houses
        self.shed_area_threshold = 5000   # Max area for a shed
        # Houses are any quadrilaterals with area greater than shed_area_threshold and less than max_area

    def get_meters_per_pixel(self, latitude, zoom_level):
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
            time.sleep(2)  # Wait for the page to load
        except Exception as e:
            self.log(f"Error initializing WebDriver: {e}")
            self.show_popup("Error", f"Error initializing WebDriver: {e}")

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
            self.log(f"Error in hide_elements: {e}")

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
            self.mark_houses(csv_folder, excel_folder, save_path)
        except Exception as e:
            self.log(f"Error processing screenshot: {e}")
            self.show_popup("Error", f"Error processing screenshot: {e}")
        finally:
            self.rubber_band.hide()

    def perform_corner_actions(self, rect):
        pyautogui.sleep(1)

        try:
            # Top-left corner
            pyautogui.moveTo(rect.left(), rect.top())
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.top_left_coords = self.get_coords_from_maps()
            self.log(f"Top-left corner coordinates: {self.top_left_coords}")

            # Bottom-right corner
            pyautogui.moveTo(rect.right(), rect.bottom())
            pyautogui.click(button='right')
            pyautogui.sleep(1)
            self.bottom_right_coords = self.get_coords_from_maps()
            self.log(f"Bottom-right corner coordinates: {self.bottom_right_coords}")

        except Exception as e:
            self.log(f"Error performing corner actions: {e}")
            self.show_popup("Error", f"Error obtaining coordinates: {e}")

    def get_coords_from_maps(self):
        try:
            element = self.driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[1]/div/div')
            coords = element.text
            return coords
        except Exception as e:
            self.log(f"Error getting coordinates: {e}")
            return ""

    def calculate_distance(self, lat1, lon1, lat2, lon2):
        """
        Вычисление расстояния между двумя координатами (в метрах) с использованием формулы Haversine.
        """
        R = 6371000  # Радиус Земли в метрах
        lat1_rad = math.radians(lat1)
        lon1_rad = math.radians(lon1)
        lat2_rad = math.radians(lat2)
        lon2_rad = math.radians(lon2)

        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad

        a = math.sin(dlat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2) ** 2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

        return R * c  # Возвращает расстояние в метрах

    def merge_close_houses(self, house_coordinates, merge_distance):
        """
        Groups houses that are close to each other and returns the center coordinates.
        """
        merged_houses = []
        used = set()

        for i, (lat1, lon1) in enumerate(house_coordinates):
            if i in used:
                continue

            close_houses = [(lat1, lon1)]
            used.add(i)

            for j, (lat2, lon2) in enumerate(house_coordinates):
                if j != i and j not in used:
                    distance = self.calculate_distance(lat1, lon1, lat2, lon2)
                    if distance <= merge_distance:
                        close_houses.append((lat2, lon2))
                        used.add(j)

            if close_houses:
                avg_lat = sum([lat for lat, lon in close_houses]) / len(close_houses)
                avg_lon = sum([lon for lat, lon in close_houses]) / len(close_houses)
                merged_houses.append((avg_lat, avg_lon))

        return merged_houses

    def detect_houses_on_map(self, image_path, latitude, longitude, meters_per_pixel):
        """
        Process the image to detect houses and mark them with improved sensitivity.
        """
        image = cv2.imread(image_path)
        if image is None:
            self.log(f"Error loading image {image_path}")
            return []

        height, width, _ = image.shape

        # Debug logging for image dimensions
        self.log(f"Image loaded with dimensions: {width}x{height}")

        # Save the original image for verification
        original_image_path = image_path.replace(".png", "_step1_original.png")
        cv2.imwrite(original_image_path, image)

        # Define color tolerance
        tolerance = 10  # Increased tolerance for better detection
        lower_bound = np.clip(self.target_color - tolerance, 0, 255)
        upper_bound = np.clip(self.target_color + tolerance, 0, 255)

        # Create a mask for the target color
        mask = cv2.inRange(image, lower_bound, upper_bound)
        self.log(f"Mask created with bounds: lower={lower_bound}, upper={upper_bound}")

        # Apply morphological operations to clean the mask
        kernel_size = 3  # Adjust if needed for sensitivity
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
        opened_mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
        closed_mask = cv2.morphologyEx(opened_mask, cv2.MORPH_CLOSE, kernel)

        # Find contours in the mask
        contours, _ = cv2.findContours(closed_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        self.log(f"Contours detected: {len(contours)}")

        house_coordinates = []

        # Thresholds for filtering
        small_house_threshold = 200  # Reduced threshold to capture smaller houses
        large_house_threshold = 1500  # Threshold for residential houses

        # For visualization
        houses_colored_image = np.zeros_like(image)

        for contour in contours:
            epsilon = 0.02 * cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, epsilon, True)

            if len(approx) >= 4:  # Include polygons with 4 or more vertices
                area = cv2.contourArea(approx)

                # Logging for each detected contour's area
                self.log(f"Contour area: {area}")

                # Filtering based on area
                if area > small_house_threshold:
                    if area >= large_house_threshold:
                        # Mark house in red
                        cv2.drawContours(houses_colored_image, [contour], -1, (0, 0, 255), cv2.FILLED)

                        # Calculate the center of the house
                        M = cv2.moments(contour)
                        if M["m00"] != 0:
                            cx = int(M["m10"] / M["m00"])
                            cy = int(M["m01"] / M["m00"])

                            dx_pixels = cx - (width / 2)
                            dy_pixels = cy - (height / 2)

                            dx_meters = dx_pixels * meters_per_pixel
                            dy_meters = dy_pixels * meters_per_pixel

                            delta_lat = -dy_meters / 110540
                            delta_lon = dx_meters / (111320 * math.cos(math.radians(latitude)))

                            house_lat = latitude + delta_lat
                            house_lon = longitude + delta_lon

                            # Add house coordinates to the list
                            house_coordinates.append((house_lat, house_lon))
                            self.log(f"Marked house at {house_lat}, {house_lon}")
                    else:
                        # Mark house in green (smaller than large threshold but larger than shed)
                        cv2.drawContours(houses_colored_image, [contour], -1, (0, 255, 0), cv2.FILLED)
                else:
                    self.log(f"Excluded house with area: {area} (too small)")
            else:
                self.log(f"Excluded contour with {len(approx)} vertices (less than 4)")

        # Save the image with marked houses
        colored_houses_path = image_path.replace(".png", "_colored_houses_filtered.png")
        cv2.imwrite(colored_houses_path, houses_colored_image)

        self.log(f"Total houses marked: {len(house_coordinates)}")

        return house_coordinates

    def is_valid_house(self, lat, lon):
        """
        Additional check to determine if the object is a house.
        You can add extra criteria here if necessary.
        Currently, it just checks that the coordinates are unique.
        """
        # You can add a check for similar coordinates in already detected houses
        for existing_lat, existing_lon in self.house_coordinates:
            distance = self.calculate_distance(lat, lon, existing_lat, existing_lon)
            if distance < 1.5:  # Acceptable distance in meters
                return False
        return True

    def capture_all_houses(self, save_path):
        """
        Traverses the entire specified area, takes screenshots, and detects houses.
        Увеличение перекрытия между снимками для захвата всех домов.
        """
        try:
            self.house_coordinates = []
            visited_positions = set()

            top_left_lat, top_left_lon = self.parse_coords(self.top_left_coords)
            bottom_right_lat, bottom_right_lon = self.parse_coords(self.bottom_right_coords)

            if top_left_lat is None or bottom_right_lat is None:
                self.log("Invalid coordinate format. Aborting capture.")
                return

            screen_width, screen_height = pyautogui.size()

            # Увеличение перекрытия до 20%
            overlap_ratio = 0.1  # 20% overlap

            current_lat = top_left_lat

            while current_lat > bottom_right_lat:
                current_lon = top_left_lon
                meters_per_pixel = self.get_meters_per_pixel(current_lat, self.current_zoom)

                # Уменьшенный шаг с увеличенным перекрытием
                step_lon = (meters_per_pixel * screen_width / (111320 * math.cos(math.radians(current_lat)))) * (
                            1 - overlap_ratio)
                step_lat = (meters_per_pixel * screen_height / 110540) * (1 - overlap_ratio)

                while current_lon < bottom_right_lon:
                    # Проверка, не был ли уже посещён этот участок
                    position_key = (round(current_lat, 7), round(current_lon, 7))
                    if position_key not in visited_positions:
                        self.move_to_and_capture(current_lat, current_lon, save_path)
                        visited_positions.add(position_key)
                    else:
                        self.log(f"Skipping duplicate position at latitude: {current_lat}, longitude: {current_lon}")
                    current_lon += step_lon

                current_lat -= step_lat

        except Exception as e:
            self.log(f"Error capturing houses: {e}")
            self.show_popup("Error", f"Error capturing houses: {e}")

    def move_to_and_capture(self, latitude, longitude, save_path):
        """
        Перемещение по карте и захват скриншотов с последующим обнаружением домов.
        """
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

            # Детектирование домов на карте
            house_coords_on_image = self.detect_houses_on_map(
                filename, latitude, longitude, meters_per_pixel
            )

            for house_lat, house_lon in house_coords_on_image:
                # Проверяем на дубликаты перед добавлением в список
                if not self.is_duplicate(house_lat, house_lon, self.house_coordinates, meters_per_pixel):
                    self.house_coordinates.append((house_lat, house_lon))
                    self.log(f"Added house at {house_lat}, {house_lon}")
                else:
                    self.log(f"Skipped duplicate house at coordinates: {house_lat}, {house_lon}")

        except Exception as e:
            self.log(f"Error taking screenshot: {e}")
            self.show_popup("Error", f"Error taking screenshot: {e}")
    def is_duplicate(self, new_lat, new_lon, house_coordinates, meters_per_pixel):
        """
        Проверка на дубликаты координат домов.
        Если расстояние до уже сохранённого дома меньше допустимого (1 метр), то считаем его дубликатом.
        """
        allowed_distance = 3.0  # Установим допустимое расстояние для дубликатов в 2 метра

        for lat, lon in house_coordinates:
            distance = self.calculate_distance(new_lat, new_lon, lat, lon)
            if distance < allowed_distance:
                return True

        return False

    def mark_houses(self, csv_folder, excel_folder, save_path):
        try:
            village_name = self.village_input.text().strip()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Convert house coordinates to NumPy array for clustering
            coords = np.array(self.house_coordinates)

            if len(coords) == 0:
                self.log("No houses detected to mark.")
                self.show_completion_dialog(
                    "Process Completed",
                    f"No houses detected.Documents\\MapsMarkup\\{self.prefix_input.text().strip()}\\{village_name}"
                )
                return

            # Perform DBSCAN clustering
            # Epsilon is set to 5 meters; adjust if necessary
            epsilon = 5 / 6371000  # Convert 5 meters to radians

            # Convert latitude and longitude to radians for geospatial clustering
            coords_rad = np.radians(coords)

            # Define DBSCAN with Haversine metric
            db = DBSCAN(eps=epsilon, min_samples=1, algorithm='ball_tree', metric='haversine').fit(coords_rad)
            labels = db.labels_

            unique_labels = set(labels)
            clustered_coords = []

            for label in unique_labels:
                class_member_mask = (labels == label)
                cluster = coords[class_member_mask]
                centroid = cluster.mean(axis=0)
                clustered_coords.append((centroid[0], centroid[1]))

            self.log(
                f"Clustering reduced {len(self.house_coordinates)} houses to {len(clustered_coords)} unique houses.")

            # Sort coordinates by latitude (LAT)
            sorted_coordinates = sorted(clustered_coords, key=lambda coord: coord[0])

            csv_filename = os.path.join(csv_folder, f'{village_name}_{timestamp}.csv')
            with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['Marker', 'Latitude', 'Longitude']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

                # Assign markers to sorted coordinates
                for idx, (lat, lon) in enumerate(sorted_coordinates, start=1):
                    marker = f"{self.prefix_input.text().strip()}{idx:03d}"
                    writer.writerow({
                        'Marker': marker,
                        'Latitude': f"{lat:.14f}",
                        'Longitude': f"{lon:.14f}"
                    })

            # Read CSV file and convert to Excel
            df = pd.read_csv(csv_filename)
            excel_filename = os.path.join(excel_folder, f'{village_name}_{timestamp}.xlsx')
            df.to_excel(excel_filename, index=False)

            self.log(f"Houses marked and saved in {csv_filename} and {excel_filename}")

            # Display completion dialog
            self.show_completion_dialog(
                "Process Completed",
                f"Process completed.Documents\\MapsMarkup\\{self.prefix_input.text().strip()}\\{village_name}"
            )

        except Exception as e:
            self.log(f"Error marking houses: {e}")
            self.show_popup("Error", f"Error saving houses: {e}")

    def delete_images(self, directory):
        try:
            for filename in os.listdir(directory):
                if filename.endswith('.png') or filename.endswith('.jpg') or filename.endswith('.jpeg'):
                    file_path = os.path.join(directory, filename)
                    os.remove(file_path)
                    self.log(f"Deleted image file: {file_path}")
        except Exception as e:
            self.log(f"Error deleting images: {e}")
            self.show_popup("Error", f"Error deleting images: {e}")

    def parse_coords(self, coords_str):
        try:
            lat, lon = map(float, coords_str.split(','))
            return lat, lon
        except ValueError:
            self.log(f"Invalid coordinate format: {coords_str}")
            return None, None


class MainApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Screenshot Application")
        self.setGeometry(100, 100, 300, 200)

        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)

        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("Enter Code")

        self.village_input = QLineEdit()
        self.village_input.setPlaceholderText("Enter Village Name")

        self.submit_button = QPushButton("Start")
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
