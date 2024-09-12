import time
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class ZoomApp:
    def __init__(self, max_zoom_level=20):
        self.driver = None
        self.max_zoom_level = max_zoom_level
        self.setup_driver()

    def setup_driver(self):
        options = Options()
        options.add_experimental_option("debuggerAddress", "localhost:9222")
        try:
            self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=options)
        except Exception as e:
            self.log(f"Error initializing WebDriver: {e}")
            self.show_popup("Error", f"Error initializing WebDriver: {e}")

    def show_popup(self, title, message):
        # Реализуйте метод для отображения всплывающих сообщений
        print(f"{title}: {message}")

    def log(self, message):
        # Реализуйте метод для логирования сообщений
        print(message)

    def get_zoom_level(self):
        try:
            zoom_script = "return window.map.getZoom();"
            zoom_level = self.driver.execute_script(zoom_script)
            return zoom_level
        except Exception as e:
            self.log(f"Error getting zoom level: {e}")
            return None

    def zoom_in(self):
        current_zoom_level = self.get_zoom_level()
        if current_zoom_level is None:
            self.log("Failed to get the current zoom level.")
            return

        if current_zoom_level >= self.max_zoom_level:
            self.log("Maximum zoom level reached.")
            self.show_popup("Zoom Limit", "Maximum zoom level reached.")
            return

        try:
            zoom_script = "document.querySelector('button[aria-label=\"Zoom in\"]').click();"
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="Zoom in"]'))
            )
            self.driver.execute_script(zoom_script)
            time.sleep(2)  # Подождите, пока зум выполнится
            self.log("Zoomed in successfully.")
        except Exception as e:
            self.log(f"Error during zooming in: {e}")
            self.show_popup("Error", f"Error during zooming in: {e}")

    def close_driver(self):
        if self.driver:
            self.driver.quit()

if __name__ == "__main__":
    app = ZoomApp(max_zoom_level=20)  # Установите желаемый уровень максимального зума
    app.zoom_in()
    app.close_driver()
