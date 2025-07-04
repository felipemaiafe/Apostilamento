import os
import sys
import re
import logging
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException

def get_base_path():
    """Gets the base path, accounting for PyInstaller's temporary directory."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # If not bundled, use the script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path

BASE_PATH = get_base_path()

# Constants
FOLDER_ADM = os.path.join(BASE_PATH, "DIARIOS_E_DITAIS", "ADM")
FOLDER_PROF = os.path.join(BASE_PATH, "DIARIOS_E_DITAIS")
MAX_RETRIES = 3
RETRY_DELAY = 2
MAX_ATTEMPTS = 2
DOCUMENT_TREE_REFRESH_DELAY = 10

def automate_Edital(driver, year_to_find, cargo_text, current_date, process_xpath, callbacks):
    """Automates Edital document creation and verification with retry logic"""
    is_administrativo = bool(re.search(r"Admin?istrativo|Analista|Agente.*Administrativo", cargo_text, re.IGNORECASE))

    def switch_frame(xpath, reset_to_default=True):
        """Switch to a specific frame with improved handling"""
        if reset_to_default:
            driver.switch_to.default_content()  # Reset context before switching
        try:
            # First find the element
            frame_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            # Then switch to it
            driver.switch_to.frame(frame_element)
            time.sleep(0.5)  # Short wait to ensure frame is fully loaded
            return True
        except TimeoutException:
            logging.error(f"Timeout: Could not find frame with XPath {xpath}")
        except NoSuchElementException:
            logging.error(f"Frame with XPath {xpath} not found")
        return False

    def click_element(xpath):
        """Click an element with retries"""
        for attempt in range(MAX_RETRIES):
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                return True
            except (NoSuchElementException, TimeoutException, StaleElementReferenceException) as e:
                logging.error(f"Attempt {attempt+1} failed to click element with XPath {xpath}: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error(f"Failed to click element with XPath {xpath} after maximum retries")
                time.sleep(RETRY_DELAY)
        return False

    def send_keys_to_element(xpath, keys):
        """Send keys to an element with retries"""
        for attempt in range(MAX_RETRIES):
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath)))
                element.clear()
                element.send_keys(keys)
                return True
            except (NoSuchElementException, TimeoutException, StaleElementReferenceException) as e:
                logging.error(f"Attempt {attempt+1} failed to send keys to element with XPath {xpath}: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error(f"Failed to send keys to element with XPath {xpath} after maximum retries")
                time.sleep(RETRY_DELAY)
        return False

    def select_dropdown_option(xpath, option_text):
        """Select an option from a dropdown with retries"""
        for attempt in range(MAX_RETRIES):
            try:
                dropdown_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath)))
                select = Select(dropdown_element)
                select.select_by_visible_text(option_text)
                return True
            except (NoSuchElementException, TimeoutException, StaleElementReferenceException) as e:
                logging.error(f"Attempt {attempt+1} failed to select option '{option_text}' from dropdown with XPath {xpath}: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error(f"Failed to select option '{option_text}' from dropdown with XPath {xpath} after maximum retries")
                time.sleep(RETRY_DELAY)
        return False

    def reset_process_state():
        """Reset the process state by clicking the process number"""
        try:
            logging.info("Resetting process state...")
            driver.switch_to.default_content()
            if not switch_frame('//*[@id="ifrArvore"]', reset_to_default=True):
                raise Exception("Failed to switch to 'ifrArvore' frame")
            if not click_element(process_xpath):
                raise Exception("Failed to click process number")
            logging.info("Process state reset successfully")
            return True
        except Exception as e:
            logging.error(f"Failed to reset process state: {str(e)}")
            return False

    def verify_document_in_tree(document_name, attempt_count=1):
        """Verify if the document exists in the tree"""
        try:
            driver.switch_to.default_content()
            if not switch_frame('//*[@id="ifrArvore"]', reset_to_default=True):
                raise Exception("Failed to switch to 'ifrArvore' frame")
            document_elements = WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]')))
            target_text = f"Edital {document_name}"
            matching_elements = [element for element in document_elements if target_text in element.text]
            if matching_elements:
                return True
            logging.warning(f"Document '{target_text}' not found in tree (attempt {attempt_count})")
            return False
        finally:
            driver.switch_to.default_content()

    def determine_document_types(year):
        """Determine which document types are available for a given year"""
        year_documents = {
            "1988": ["CAPA", "LISTA"],
            "1993": ["CAPA", "LISTA"],
            "1999": ["CAPA", "LISTA"],
            "2004": ["CAPA", "LISTA"],
            "2005": ["CAPA"],
            "2006": ["CAPA", "LISTA"],
            "2008": ["CAPA"],
            "2010": ["LISTA"]
        }
        return year_documents.get(str(year), [])

    def create_and_fill_document(document_name):
        """Create and fill the Edital document with retries"""
        for attempt in range(MAX_ATTEMPTS):
            current_attempt = attempt + 1
            try:
                if not switch_frame('//*[@id="ifrConteudoVisualizacao"]', reset_to_default=True):
                    raise Exception("Failed to switch to 'ifrConteudoVisualizacao' frame")
                
                # Step 1: Click at the "add document" button
                if not click_element('//*[@id="divArvoreAcoes"]/a[1]/img'):
                    raise Exception("Failed to click 'Incluir Documento' button")
                
                # Find the visualization frame without resetting to default content
                if not switch_frame('//*[@id="ifrVisualizacao"]', reset_to_default=False):
                    # If direct access fails, try resetting and then switching
                    if not switch_frame('//*[@id="ifrVisualizacao"]', reset_to_default=True):
                        raise Exception("Failed to switch to 'ifrVisualizacao' frame")
                
                # Step 2: Select "Externo" from a list
                if not click_element('//*[@id="tblSeries"]/tbody/tr[1]/td/a[2]'):
                    raise Exception("Failed to click 'Externo' option")
                
                # Step 3: Select "Edital" from a dropdown list
                if not select_dropdown_option('//*[@id="selSerie"]', "Edital"):
                    raise Exception("Failed to select 'Edital' from dropdown")
                
                time.sleep(4)  # Ensure page has reloaded
                
                # Step 4: Fill required fields
                if not send_keys_to_element('//*[@id="txtDataElaboracao"]', current_date):
                    raise Exception("Failed to send keys to 'Date' field")
                if not send_keys_to_element('//*[@id="txtNomeArvore"]', document_name):
                    raise Exception("Failed to send keys to 'Nome do Documento' field")
                if not click_element('//*[@id="divOptNato"]/div/label'):
                    raise Exception("Failed to click 'Nato' checkbox")
                if not click_element('//*[@id="divOptPublico"]/div/label'):
                    raise Exception("Failed to click 'Público' checkbox")
                
                # Step 5: Upload file
                file_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="filArquivo"]')))

                # Construct the correct file path based on the document type and administrative status
                if is_administrativo:
                    # For administrative documents, use the ADM folder and include ADM in the filename
                    file_path = os.path.join(FOLDER_ADM, f"Edital___{year_to_find}_ADM_{document_name.upper()}.pdf")
                else:
                    # For professor documents, use the PROF folder without ADM in the filename
                    file_path = os.path.join(FOLDER_PROF, f"Edital___{year_to_find}_{document_name.upper()}.pdf")

                # Check if the file exists
                if not os.path.exists(file_path):
                    logging.error(f"File not found: {file_path}")
                    return False
                
                file_input.send_keys(file_path)
                
                if not WebDriverWait(driver, 100).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="tblAnexos"]/tbody/tr/td[2]'))):
                    raise Exception("File not attached to document")
                
                # Step 6: Save document
                if not click_element('//*[@id="btnSalvar"]'):
                    raise Exception("Failed to click 'Salvar' button")
                
                time.sleep(DOCUMENT_TREE_REFRESH_DELAY)
                
                if verify_document_in_tree(document_name, attempt_count=current_attempt):
                    logging.info(f"{document_name} uploaded successfully")
                    return True
                logging.warning(f"Verification failed for {document_name}")
            except Exception as e:
                logging.error(f"Attempt {current_attempt} error: {str(e)}")
                if current_attempt < MAX_ATTEMPTS:
                    if not reset_process_state():
                        logging.error("Aborting retry due to failed state reset")
                        return False
        return False

    try:
        available_documents = determine_document_types(year_to_find)
        if not available_documents:
            logging.error(f"No document types defined for year {year_to_find}")
            # Mark potential items as failed
            callbacks['update_checklist']('Edital CAPA', False)
            callbacks['update_checklist']('Edital LISTA', False)
            return False

        all_success = True
        for document_type in ["CAPA", "LISTA"]: # Check against a fixed list
            if document_type in available_documents:
                logging.info(f"Processing Edital {document_type}...")
                success = create_and_fill_document(document_type)
                callbacks['update_checklist'](f'Edital {document_type}', success)
                if not success:
                    logging.error(f"Failed to upload {document_type} for year {year_to_find}")
                    all_success = False
            else:
                pass
        
        return all_success
        
    except Exception as e:
        logging.error(f"Edital processing failed with a critical error: {str(e)}")
        callbacks['update_checklist']('Edital CAPA', False)
        callbacks['update_checklist']('Edital LISTA', False)
        return False