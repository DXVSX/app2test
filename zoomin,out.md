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
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By


class ScreenshotApp(QMainWindow):
    def __init__(self, tray_icon, log_widget, prefix_input, village_input, zoom_input):
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
        self.zoom_input = zoom_input
        self.setup_driver()

        # Параметры зума
        self.min_zoom = 0
        self.max_zoom = 10  # Можно изменить пользователем в интерфейсе
        self.current_zoom = 0  # Начальный уровень зума

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
        Метод увеличения масштаба с использованием прокрутки мыши
        """
        if self.current_zoom < self.max_zoom:
            try:
                pyautogui.scroll(500)  # Прокрутка вверх для увеличения зума
                self.current_zoom += 1
                time.sleep(1)
                self.log(f"Zoom level: {self.current_zoom}")
            except Exception as e:
                self.log(f"Error during zoom in: {e}")
                self.show_popup("Error", f"Error during zoom in: {e}")

    def zoom_out(self):
        """
        Метод уменьшения масштаба (если потребуется)
        """
        if self.current_zoom > self.min_zoom:
            try:
                pyautogui.scroll(-500)  # Прокрутка вниз для уменьшения зума
                self.current_zoom -= 1
                time.sleep(1)
                self.log(f"Zoom level: {self.current_zoom}")
            except Exception as e:
                self.log(f"Error during zoom out: {e}")
                self.show_popup("Error", f"Error during zoom out: {e}")


class SystemTrayApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.main_window = QWidget()
        self.layout = QVBoxLayout()
        self.log_widget = QTextEdit()
        self.log_widget.setReadOnly(True)
        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("Введите код")
        self.village_input = QLineEdit()
        self.village_input.setPlaceholderText("Введите название села")
        self.zoom_input = QLineEdit()
        self.zoom_input.setPlaceholderText("Введите лимит для зума")

        self.main_window.setWindowTitle("Map Markup")
        self.main_window.setWindowIcon(QIcon('icon.png'))

        self.screenshot_button = QPushButton("Выбрать область для сканирования")
        self.screenshot_button.clicked.connect(self.start_screenshot)

        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.submit_clicked)

        self.layout.addWidget(self.log_widget)
        self.layout.addWidget(self.prefix_input)
        self.layout.addWidget(self.village_input)
        self.layout.addWidget(self.zoom_input)
        self.layout.addWidget(self.screenshot_button)
        self.layout.addWidget(self.submit_button)
        self.main_window.setLayout(self.layout)

        self.tray_icon = QSystemTrayIcon(QIcon('icon.png'), self.app)
        tray_menu = QMenu()
        exit_action = QAction("Exit", self.main_window)
        exit_action.triggered.connect(self.exit_app)
        tray_menu.addAction(exit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        self.screenshot_app = ScreenshotApp(self.tray_icon, self.log_widget, self.prefix_input, self.village_input, self.zoom_input)

    def start_screenshot(self):
        self.screenshot_app.start_screenshot()

    def submit_clicked(self):
        try:
            max_zoom = int(self.zoom_input.text().strip())
            self.screenshot_app.max_zoom = max_zoom
            self.show_popup("Success", f"Zoom limit set to {max_zoom}")
        except ValueError:
            self.show_popup("Error", "Please enter a valid number for the zoom limit.")

    def show_popup(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def exit_app(self):
        self.tray_icon.hide()
        sys.exit()

    def run(self):
        self.main_window.show()
        sys.exit(self.app.exec_())


if __name__ == "__main__":
    tray_app = SystemTrayApp()
    tray_app.run()
