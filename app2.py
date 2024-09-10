import tkinter as tk
from tkinter import messagebox
import webbrowser
from PIL import ImageGrab
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

class CoordinatesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Maps Coordinates")

        self.entries = []
        self.add_coordinate_entry()

        # Кнопка для открытия Google Maps
        self.submit_button = tk.Button(root, text="Open Google Maps", command=self.open_google_maps_with_locations)
        self.submit_button.grid(row=8, column=0, columnspan=2, pady=10)

        # Кнопка для создания скриншота
        self.screenshot_button = tk.Button(root, text="Take Screenshot", command=self.take_screenshot)
        self.screenshot_button.grid(row=9, column=0, columnspan=2, pady=10)

        # Автоматическое открытие Google Maps при запуске
        self.open_google_maps_with_locations()

    def add_coordinate_entry(self):
        """Добавляет новую строку для ввода координат, если их меньше 7."""
        row = len(self.entries)
        tk.Label(self.root, text=f"Coordinates {row + 1} (latitude, longitude):").grid(row=row, column=0, padx=10, pady=10)
        entry = tk.Entry(self.root, width=50)
        entry.grid(row=row, column=1, padx=10, pady=10)
        self.entries.append(entry)

        # Удаление предыдущей кнопки "Add Another Coordinate", если она существует
        if hasattr(self, 'add_button'):
            self.add_button.grid_forget()

        # Добавление новой кнопки, если есть место для добавления еще одной строки
        if len(self.entries) < 7:
            self.add_button = tk.Button(self.root, text="Add Another Coordinate", command=self.add_coordinate_entry)
            self.add_button.grid(row=row + 1, column=0, columnspan=2, pady=10)

    def setup_driver(self):
        """Настраивает WebDriver с указанным адресом отладчика."""
        options = Options()
        options.add_experimental_option("debuggerAddress", "localhost:9222")
        chromedriver_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'chromedriver.exe')
        print(f"Using chromedriver path: {chromedriver_path}")  # Вывод пути для проверки
        try:
            self.driver = webdriver.Chrome(service=Service(chromedriver_path), options=options)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize WebDriver: {str(e)}")
            print(f"Error initializing WebDriver: {str(e)}")  # Дополнительный вывод в консоль

    def open_google_maps_with_locations(self):
        try:
            coord_list = []
            for entry in self.entries:
                coords = entry.get().strip()
                if coords:
                    coord_list.append(coords)

            # Проверка корректности каждого набора координат
            for coord in coord_list:
                parts = coord.split(',')
                if len(parts) != 2:
                    raise ValueError("Each coordinate must be in the format 'latitude,longitude'.")

                latitude = parts[0].strip()
                longitude = parts[1].strip()

                try:
                    latitude = float(latitude)
                    longitude = float(longitude)
                except ValueError:
                    raise ValueError("Latitude and longitude must be valid numbers.")

            # Формирование URL для отображения на карте
            markers = '|'.join([f"{lat},{lng}" for lat, lng in (coord.split(',') for coord in coord_list)])
            url = f"https://www.google.com/maps/dir/?api=1&markers={markers}"
            webbrowser.open(url)
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))

    def take_screenshot(self):
        if len(self.entries) < 2:
            messagebox.showerror("Error", "You must enter at least 2 coordinates to take a screenshot.")
            return
        
        try:
            # Получение списка координат
            coord_list = []
            for entry in self.entries:
                coords = entry.get().strip()
                if coords:
                    coord_list.append(coords)

            # Настройка WebDriver
            self.setup_driver()

            if not hasattr(self, 'driver'):
                return

            # Открываем Google Maps с координатами
            markers = '|'.join([f"{lat},{lng}" for lat, lng in (coord.split(',') for coord in coord_list)])
            url = f"https://www.google.com/maps/dir/?api=1&markers={markers}"
            self.driver.get(url)

            # Ожидаем загрузку карты
            time.sleep(5)  # Задержка для полной загрузки карты

            # Папка для сохранения скриншота
            folder_path = os.path.join(os.getcwd(), 'MapsMarkup')
            os.makedirs(folder_path, exist_ok=True)
            
            # Полный путь для сохранения файла
            screenshot_path = os.path.join(folder_path, 'google_maps_screenshot.png')

            # Скриншот
            screenshot = self.driver.get_screenshot_as_file(screenshot_path)

            # Проверка успешного создания скриншота
            if screenshot:
                messagebox.showinfo("Screenshot", f"Screenshot saved as {screenshot_path}")
            else:
                messagebox.showerror("Error", "Failed to take screenshot.")

            # Закрываем браузер
            self.driver.quit()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to take screenshot: {str(e)}")

# Создание основного окна
root = tk.Tk()
app = CoordinatesApp(root)
root.mainloop()
