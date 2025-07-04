import glob
import os
import logging
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
from PyPDF2 import PdfMerger

# Constants
MAX_RETRIES = 3
RETRY_DELAY = 2
DOCUMENT_TREE_REFRESH_DELAY = 10

def ensure_file_saved(filepath, timeout=10):
    """Ensure the file is saved and not empty"""
    elapsed_time = 0
    while elapsed_time < timeout:
        if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
            return True
        time.sleep(1)
        elapsed_time += 1
    return False

def merge_pdfs(temp_dir_path):
    """
    Merges PDFs found in a temp directory and saves the result there.
    Returns the path to the combined PDF.
    """

    # Find all 'ficha_financeira_*.pdf' files in the directory
    pdf_files = glob.glob(os.path.join(temp_dir_path, "ficha_financeira_*.pdf"))
    
    if not pdf_files:
        logging.error(f"No Ficha Financeira PDF files found in {temp_dir_path}")
        return None
    
    # Sort files to ensure correct order (e.g., page 1, 2, 3)
    pdf_files.sort()
    
    # Define the output path for the combined file inside the same temp directory
    combined_pdf_path = os.path.join(temp_dir_path, "ficha_financeira_combined.pdf")

    merger = PdfMerger()
    try:
        for pdf in pdf_files:
            if ensure_file_saved(pdf):
                merger.append(pdf)
            else:
                logging.warning(f"Could not validate file, skipping: {pdf}")
        
        if len(merger.pages) > 0:
            merger.write(combined_pdf_path)
            logging.info("Ficha Financeira salva.")
            return combined_pdf_path
        else:
            logging.error("No valid pages were merged.")
            return None

    except Exception as e:
        logging.error(f"PDF merge failed: {str(e)}")
        return None
    finally:
        merger.close()

def switch_frame(driver, xpath, reset_to_default=True):
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

def click_element(driver, xpath):
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

def send_keys_to_element(driver, xpath, keys):
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

def select_dropdown_option(driver, xpath, option_text):
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

def verify_ficha_in_tree(driver):
    """Verify if Ficha exists in tree using Edital.py's logic"""
    try:
        if not switch_frame(driver, '//*[@id="ifrArvore"]'):
            return False
            
        document_elements = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]')))
        
        target_text = "Ficha Financeira"
        matching_elements = [element for element in document_elements if target_text in element.text]
        
        return len(matching_elements) > 0
    except Exception as e:
        logging.error(f"Verification failed: {str(e)}")
        return False
    finally:
        driver.switch_to.default_content()

def upload_Ficha_Financeira(driver, current_date, callbacks, combined_pdf_path):
    """Main upload function with aligned verification logic"""

    try:
        for attempt in range(MAX_RETRIES):
            try:
                if not switch_frame(driver, '//*[@id="ifrConteudoVisualizacao"]', reset_to_default=True):
                    raise Exception("Failed to switch to 'ifrConteudoVisualizacao' frame")
                
                # Step 1: Click at the "add document" button
                if not click_element(driver, '//*[@id="divArvoreAcoes"]/a[1]/img'):
                    raise Exception("Failed to click 'Incluir Documento' button")
                
                # Find the visualization frame without resetting to default content
                if not switch_frame(driver, '//*[@id="ifrVisualizacao"]', reset_to_default=False):
                    # If direct access fails, try resetting and then switching
                    if not switch_frame(driver, '//*[@id="ifrVisualizacao"]', reset_to_default=True):
                        raise Exception("Failed to switch to 'ifrVisualizacao' frame")
                
                # Step 2: Select "Externo" from a list
                if not click_element(driver, '//*[@id="tblSeries"]/tbody/tr[1]/td/a[2]'):
                    raise Exception("Failed to click 'Externo' option")
                
                # Step 3: Select "Ficha Financeira" from a dropdown list
                if not select_dropdown_option(driver, '//*[@id="selSerie"]', "Ficha Financeira"):
                    raise Exception("Failed to select 'Ficha Financeira' from dropdown")
                
                time.sleep(4)  # Ensure page has reloaded
                
                # Fill form fields
                if not send_keys_to_element(driver, '//*[@id="txtDataElaboracao"]', current_date):
                    raise Exception("Failed to send keys to 'Date' field")
                if not click_element(driver, '//*[@id="divOptNato"]/div/label'):
                    raise Exception("Failed to click 'Nato' checkbox")
                if not click_element(driver, '//*[@id="divOptPublico"]/div/label'):
                    raise Exception("Failed to click 'Público' checkbox")
                
                # File upload
                if not os.path.exists(combined_pdf_path):
                    logging.error(f"Combined PDF not found at path: {combined_pdf_path}")
                    callbacks['update_checklist']('Ficha Financeira', False)
                    return False
                
                file_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="filArquivo"]'))
                )
                
                # Send the file path to the input
                file_input.send_keys(combined_pdf_path)
                
                WebDriverWait(driver, 100).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="tblAnexos"]/tbody/tr/td[2]')))
                
                click_element(driver, '//*[@id="btnSalvar"]')
                time.sleep(DOCUMENT_TREE_REFRESH_DELAY)
                
                # Enhanced verification loop
                for check in range(3):
                    if verify_ficha_in_tree(driver):
                        logging.info("Ficha Financeira uploaded successfully")
                        callbacks['update_checklist']('Ficha Financeira', True)
                        return True
                    logging.warning(f"Verification retry {check+1}/3")
                    time.sleep(3)
                
                logging.error("Final verification failed")
                # Fall through to the outer except block or return False
                
            except Exception as e:
                logging.error(f"Attempt {attempt+1} error: {str(e)}")
                if attempt == MAX_RETRIES-1:
                    logging.error("Max retries reached for Ficha Financeira upload.")
                    callbacks['update_checklist']('Ficha Financeira', False)
                    return False
                time.sleep(RETRY_DELAY)
    
    except Exception as final_e:
        # Catch any other unexpected error outside the loop
        logging.error(f"A critical error occurred in upload_Ficha_Financeira: {final_e}")
    
    # If the function exits without returning True, it's a failure.
    callbacks['update_checklist']('Ficha Financeira', False)
    return False
    
##################### TEST CODE #####################

if __name__ == "__main__":
    merge_pdfs()