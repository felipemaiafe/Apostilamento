import logging
import os
import sys

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

logging.getLogger('WDM').setLevel(logging.WARNING)

if getattr(sys, 'frozen', False):
    BASE_PATH_FOR_SAVING = os.path.dirname(sys.executable)
else:
    BASE_PATH_FOR_SAVING = os.path.dirname(os.path.abspath(__file__))

FAILED_PROCESSES_FILE = os.path.join(BASE_PATH_FOR_SAVING, "failed_processes.txt")
SUCCESSFUL_PROCESSES_FILE = os.path.join(BASE_PATH_FOR_SAVING, "successful_processes.txt")

def start_new_driver_session(download_dir=None):
    """
    Starts a new Selenium WebDriver session with automatic ChromeDriver management.
    
    Args:
        download_dir (str, optional): The absolute path for the download directory. 
                                      Defaults to None, which uses the browser's default.
    
    Returns:
        webdriver.Chrome: The configured WebDriver instance.
    """

    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    prefs = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # Direct download of PDFs
    }
    
    if download_dir:
        prefs["savefile.default_directory"] = download_dir
        
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--kiosk-printing")  # Bypass print preview if needed

    service = ChromeService(ChromeDriverManager().install())
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()

    return driver

def load_failed_processes():
    """Load failed process numbers from the .txt file."""
    try:
        with open(FAILED_PROCESSES_FILE, "r") as f:
            return set(line.strip() for line in f)
    except FileNotFoundError:
        return set()

def save_failed_process(process_number):
    """Append a failed process number to the .txt file."""
    with open(FAILED_PROCESSES_FILE, "a") as f:
        f.write(f"{process_number}\n")
        logging.info(f"Process {process_number} added to failed processes.")

def load_successful_processes():
    """Load successful process numbers from the .txt file."""
    try:
        with open(SUCCESSFUL_PROCESSES_FILE, "r") as f:
            return set(line.strip() for line in f)
    except FileNotFoundError:
        return set()

def save_successful_process(process_number):
    """Append a successful process number to the .txt file."""
    with open(SUCCESSFUL_PROCESSES_FILE, "a") as f:
        f.write(f"{process_number}\n")
        logging.info(f"Process {process_number} added to successful processes.")