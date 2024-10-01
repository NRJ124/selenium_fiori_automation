from config.config import *
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import logging
import tkinter as tk
from tkinter import messagebox
import ctypes


logger = logging.getLogger(__name__)

def start_selenium():
    edge_options = Options()
    edge_options.add_argument('--headless')  # Enable headless mode
    edge_options.add_argument('--disable-gpu')  # Disable GPU for better compatibility
    edge_options.add_experimental_option("detach", True)
    edge_options.add_argument("--start-maximized")
    service = Service(webdriver_path)
    driver = webdriver.Edge(service=service, options=edge_options)

    # Open URL
    try:
        driver.get(URL)
        WebDriverWait(driver, 30).until(EC.title_contains("Intercompany"))
        logger.info("Successfully loaded Fiori page.")
    except TimeoutException:
        logger.error("Loading took too much time!")
    except WebDriverException as e:
        logger.error(f"WebDriver error: {e}")

    return driver

def show_popup():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    messagebox.showinfo("Notification", "Your script has completed successfully!")
    root.destroy()  # Close the tkinter instance

def prevent_sleep():
    # Prevents the system from sleeping or turning off the screen
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)

def restore_sleep():
    # Allows the system to sleep again
    ES_CONTINUOUS = 0x80000000
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)




