import os
import re
import time
import fitz  # PyMuPDF
import logging
import tempfile
import shutil

from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from RHnet import automate_RHnet
from Edital import automate_Edital
from Apostila import automate_Apostila
from Despacho import automate_Despacho
from Ficha_Financeira import merge_pdfs, upload_Ficha_Financeira
from utils import save_failed_process, save_successful_process

# Constants
URL_SEI = "https://sei.go.gov.br"

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class StopRequestException(Exception):
    """Custom exception to signal a graceful stop requested by the user."""
    pass

def check_for_stop_and_pause(stop_event, pause_event):
    """Checks for stop or pause events and acts accordingly."""
    if stop_event.is_set():
        logging.warning("Stop request detected. Halting workflow.")
        raise StopRequestException("Stop requested by user.")
    
    # Check for pause
    if pause_event.is_set():
        logging.info("|| Automation Paused. Waiting for Resume... ||")
        while pause_event.is_set():
            if stop_event.is_set():
                logging.warning("Stop request detected during pause. Halting workflow.")
                raise StopRequestException("Stop requested by user during pause.")
            time.sleep(1) 

def login_to_system(driver, username, password):
    """Log in to the SEI system"""
    try:
        driver.get(URL_SEI)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtUsuario"]'))).send_keys(username)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pwdSenha"]'))).send_keys(password)
        dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selOrgao"]')))
        Select(dropdown_element).select_by_visible_text("SEDUC")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sbmAcessar"]'))).click()

        # Handle pop-up
        try:
            close_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[starts-with(@id, 'divInfraSparklingModalClose')]//img[@title='Fechar janela (ESC)']"))
            )
            close_button.click()
        except TimeoutException:
            logging.warning("Pop-up message did not appear.")

    except Exception as e:
        logging.error(f"Error during login: {e}")
        return False
    
    return True

def click_element(driver, xpath, retries=3):
    """Click an element with retries"""
    for _ in range(retries):
        try:
            element = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            return True
        except (NoSuchElementException, TimeoutException):
            logging.error(f"Could not click element with xpath: {xpath}")
    return False

def initial_navigate_and_filter(driver):
    """Navigates to process view and clicks the 'Ver atribuídos a mim' filter."""
    try:
        driver.switch_to.default_content()
        # Define XPaths
        controle_button_xpath = "//img[contains(@src, 'controle_processos_barra.svg')]"
        process_list_table_xpath = '/html/body/div[1]/div/div[2]/form/div/div[5]/div[2]/div/table/tbody'
        filter_link_xpath = "//a[normalize-space()='Ver atribuídos a mim']"

        logging.info("Procurando processos...")

        # 1. Go to the main process view first using the top button
        if not click_element(driver, controle_button_xpath):
            logging.error("Failed to click 'Controle de Processos' button for initial view.")
            return False
        time.sleep(1)

        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, process_list_table_xpath))
            )
        except TimeoutException:
            logging.error("Timeout waiting for process list table to appear AFTER clicking 'Controle de Processos'. Cannot proceed to filter.")
            return False

        # 2. Now click the "Ver atribuídos a mim" filter link
        if not click_element(driver, filter_link_xpath):
            logging.error("Failed to click 'Ver atribuídos a mim' filter link.")
            return False

        # 3. Wait for the list to reload/filter after the click
        time.sleep(1)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, process_list_table_xpath))
        )
        return True

    except Exception as e:
        logging.error(f"Failed during initial navigation and filtering: {e}")
        return False

def process_navigation(driver, failed_processes, successful_processes, stop_event, pause_event):
    """Navigate through processes and select a valid one"""

    try:
        driver.switch_to.default_content()
    except Exception as sw_err:
        logging.warning(f"Could not switch to default content before table search: {sw_err}")

    while True:
        check_for_stop_and_pause(stop_event, pause_event)
        try:
            process_list_table_xpath = '/html/body/div[1]/div/div[2]/form/div/div[5]/div[2]/div/table/tbody'
            table_body = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH,  process_list_table_xpath))
            )
            rows = table_body.find_elements(By.TAG_NAME, 'tr')
        except (NoSuchElementException, TimeoutException):
            logging.error("Process list table could not be loaded or found.")
            return None

        for i in range(len(rows) - 1, -1, -1):
            current_row = rows[i]
            try:
                process_number = extract_process_number(current_row)
                if process_number is None:
                    continue

                if process_number in failed_processes or process_number in successful_processes:
                    continue

                if not validate_white_marker(current_row):
                    logging.info(f"Process {process_number} does not have required marker. Adding to failed.")
                    failed_processes.add(process_number)
                    save_failed_process(process_number)
                    continue

                # Click the process link
                process_link_xpath = './/a[contains(@class, "processoVisualizado")]'
                process_link = WebDriverWait(current_row, 5).until(
                    EC.element_to_be_clickable((By.XPATH, process_link_xpath))
                )
                process_link.click()
                time.sleep(2)
                return process_number

            except Exception as row_e:
                 logging.error(f"Error processing row {i}: {row_e}")
                 # Continue to the next row

        logging.info("No suitable process found on this page; checking for next page.")
        try:
            next_page_xpath = '//*[@id="lnkDetalhadoProximaPaginaSuperior"]/img'
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, next_page_xpath)))
            if not click_element(driver, next_page_xpath):
                logging.error("Next page button exists but click failed. Stopping navigation.")
                return None
            logging.info("Clicked next page button.")
            check_for_stop_and_pause(stop_event, pause_event)
            time.sleep(3)
        except TimeoutException:
            # Button doesn't exist or isn't found quickly, no more pages
            logging.info("No 'next page' button found. All processes checked.")
            return False 

def return_to_filtered_list_view(driver):
    """Clicks the 'Controle de Processos' button, pauses, and waits for the list page."""
    try:
        driver.switch_to.default_content()
        controle_button_xpath = "//img[contains(@src, 'controle_processos_barra.svg')]"
        process_list_table_xpath = '/html/body/div[1]/div/div[2]/form/div/div[5]/div[2]/div/table/tbody'
        
        if not click_element(driver, controle_button_xpath):
            logging.error("Failed to click 'Controle de Processos' button.")
            return False

        time.sleep(2)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, process_list_table_xpath))
        )
        return True
    except Exception as e:
        logging.error(f"Failed to return to process list using top button: {e}")
        return False

def main_workflow(driver, process_number, failed_processes, successful_processes, callbacks, credentials, stop_event, pause_event):
    """
    Main workflow to automate the entire process, now integrated with GUI callbacks.
    
    This function orchestrates the calls to different automation modules.
    Each module is responsible for updating its own status on the GUI checklist.
    If any step fails, it raises an exception to halt the workflow for the current process.
    """
    
    current_date = datetime.now().strftime("%d/%m/%Y")
    logging.info(f"DATA: {current_date}")

    ficha_temp_dir = None
    
    try:
        # Step 1: Prerequisite Check - Extract Info from "Despacho do Gabinete"
        number_after_despacho, relevant_title, relevant_title2, chunk_of_text, cpf_number, number_in_chunk = open_and_check_despachoGAB(driver, process_number, failed_processes, successful_processes, stop_event, pause_event)
        if not all([number_after_despacho, relevant_title, relevant_title2, chunk_of_text, cpf_number]):
            raise Exception(f"Initial document check/data extraction failed for process {process_number}.")
        check_for_stop_and_pause(stop_event, pause_event)

        # Step 2: Prerequisite - Get Data from RHnet
        person_name, vinculo_number, year, cargo, ficha_temp_dir = automate_RHnet(
            cpf_number, credentials['rhnet_user'], credentials['rhnet_pass']
        )
        if not all([person_name, vinculo_number, year, cargo, ficha_temp_dir]):
            raise Exception("Failed to retrieve complete data and files from RHnet.")
        check_for_stop_and_pause(stop_event, pause_event)
        
        # Step 3: Prerequisite - Merge PDFs downloaded from RHnet
        combined_pdf_path = merge_pdfs(ficha_temp_dir)
        if not combined_pdf_path:
            raise Exception("Failed to merge Ficha Financeira PDFs.")
        check_for_stop_and_pause(stop_event, pause_event)

        # Step 4: Log key information
        logging.info(f"-----------------------")
        logging.info(f"NOME: {person_name}")
        logging.info(f"CPF: {cpf_number}")
        logging.info(f"CARGO: {cargo}")
        logging.info(f"YEAR: {year}")
        logging.info(f"-----------------------")
        check_for_stop_and_pause(stop_event, pause_event)

        # Step 5: Automate Edital
        year_to_find = determine_year_range(year)
        if year_to_find:
            process_xpath = f"//span[text()='{process_number}']/ancestor::a"
            edital_success = automate_Edital(
                driver=driver,
                year_to_find=year_to_find,
                cargo_text=cargo,
                current_date=current_date,
                process_xpath=process_xpath,
                callbacks=callbacks  # Pass callbacks down
            )
            if not edital_success:
                raise Exception("Edital processing failed.")
        else:
            logging.info("No applicable Edital year found. Skipping Edital step.")
        check_for_stop_and_pause(stop_event, pause_event)
            
        # Step 6: Check for supporting documents (Portaria, Diário)
        number_after_portaria = check_for_portaria(driver, process_number, failed_processes)
        if not number_after_portaria:
            raise Exception("'PORTARIA - SEI' not found in the document.")
        check_for_stop_and_pause(stop_event, pause_event)
        
        diario_date = check_diario_date(driver, process_number)
        if not diario_date:
            raise Exception("Diário Oficial date not found.")
        check_for_stop_and_pause(stop_event, pause_event)
        
        # Step 7: Upload Ficha Financeira
        ficha_financeira_success = upload_Ficha_Financeira(
            driver=driver,
            current_date=current_date,
            callbacks=callbacks,
            combined_pdf_path=combined_pdf_path # Pass the path
        )
        if not ficha_financeira_success:
            raise Exception("Ficha Financeira upload failed.")
        check_for_stop_and_pause(stop_event, pause_event)
        
        # Step 8: Automate Apostila
        apostila_success = automate_Apostila(
            driver, relevant_title2, number_after_portaria, process_number,
            person_name, cpf_number, chunk_of_text, relevant_title,
            number_after_despacho, vinculo_number, diario_date, number_in_chunk,
            callbacks=callbacks # Pass callbacks down
        )
        if not apostila_success:
            raise Exception("Apostila processing failed.")
        check_for_stop_and_pause(stop_event, pause_event)
        
        # Step 9: Automate Despacho
        despacho_success = automate_Despacho(
            driver=driver,
            cpf_number=cpf_number,
            process_number=process_number,
            callbacks=callbacks # Pass callbacks down
        )
        if not despacho_success:
            raise Exception("Despacho processing failed.")
        check_for_stop_and_pause(stop_event, pause_event)
            
        # Step 10: Finalization
        remove_marker_and_save(driver, process_number)
        check_for_stop_and_pause(stop_event, pause_event)
        
        # If we reach this point, the entire workflow for this process was a success.
        successful_processes.add(process_number)
        save_successful_process(process_number)

    except Exception as e:
        if type(e).__name__ == 'StopRequestException':
            logging.info(f"Análise do processo {process_number} interrompida pelo usuário.")
            raise        
        else:
            logging.error(f"Análise do processo {process_number} interrompida por um erro: {str(e)}")
            failed_processes.add(process_number)
            save_failed_process(process_number)

    finally:
        if ficha_temp_dir and os.path.exists(ficha_temp_dir):
            try:
                shutil.rmtree(ficha_temp_dir)
            except Exception as cleanup_e:
                logging.error(f"Failed to clean up temp directory {ficha_temp_dir}. Error: {cleanup_e}")
        
def extract_process_number(row):
    """Extract process number from a row"""
    try:
        process_number = row.find_element(By.CLASS_NAME, 'processoVisualizado').text.strip()
        return process_number
    except NoSuchElementException:
        return None

def validate_white_marker(row):
    """
    Validates the presence of the white marker ('marcador_branco.svg') that
    specifically has the 'APOSTILAMENTO' label. Allows for other markers to be present.
    """

    try:
        marker_xpath = "./td[2]/a[contains(@aria-label, 'APOSTILAMENTO')]/img[contains(@src, 'marcador_branco.svg')]"
        markers = row.find_elements(By.XPATH, marker_xpath)
        return len(markers) > 0
            
    except Exception as e:
        logging.error(f"An unexpected error occurred in validate_white_marker: {e}")
        return False

def locate_and_expand_tree(driver):
    """Locate and expand the document tree"""
    try:
        driver.switch_to.default_content()
        tree_iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ifrArvore"]')))
        driver.switch_to.frame(tree_iframe)
        try:
            plus_button = driver.find_element(By.XPATH, '//img[contains(@src, "mais.svg")]')
            if plus_button.is_displayed() and plus_button.is_enabled():
                plus_button.click()
                time.sleep(2)
        except NoSuchElementException:
            pass
        return True
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Error locating or expanding document tree: {e}")
        return False

def open_and_check_despachoGAB(driver, process_number, failed_processes, successful_processes, stop_event, pause_event):
    """Open and check Despacho do Gabinete document"""
    if not locate_and_expand_tree(driver):
        logging.error("Failed to locate and expand the document tree.")
        failed_processes.add(process_number)
        return None, None, None, None, None, None
    try:
        document_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]'))
        )
        despacho_element = None
        number_after_despacho = None
        for element in reversed(document_elements):
            if "Despacho do Gabinete Nº Manual" in element.text:
                despacho_text = element.text
                match = re.search(r'\((\d+)\)', despacho_text)
                if match:
                    number_after_despacho = match.group(1)
                    logging.info(f"DESPACHO GAB - SEI: {number_after_despacho}")
                else:
                    logging.warning("No number found in 'Despacho do Gabinete Nº Manual'.")

                # Scroll the element into view
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(1)

                try:
                    # Try regular click first
                    element.click()
                except Exception as e:
                    logging.warning(f"Regular click failed: {str(e)}. Trying JavaScript click.")
                    # Fallback to JavaScript click
                    driver.execute_script("arguments[0].click();", element)
                
                despacho_element = element
                break
        if not despacho_element:
            logging.error("'Despacho do Gabinete Nº Manual' not found in the document tree.")
            failed_processes.add(process_number)
            return None, None, None, None, None, None
        check_for_stop_and_pause(stop_event, pause_event)
        
        time.sleep(2)
        
        driver.switch_to.default_content()
        parent_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrConteudoVisualizacao"]'))
        )
        driver.switch_to.frame(parent_iframe)
        document_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrVisualizacao"]'))
        )
        driver.switch_to.frame(document_iframe)
        document_body = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body'))
        )
        document_text = document_body.text

        # Check for "resolvem retificar" early to avoid unnecessary processing
        if "resolvem retificar" in document_text.lower():
            logging.info("Phrase 'resolvem retificar' found.")
            remove_marker_and_save(driver, process_number)
            add_marker_and_save(driver, process_number, failed_processes, successful_processes)            
            return number_after_despacho, None, None, None, None, None
        
        despacho_match = re.search(r"DESPACHO Nº\s*([\s\S]*?)\n", document_text)
        relevant_title = despacho_match.group(0).strip() if despacho_match else "Unknown"
        relevant_title = re.sub(r"^DESPACHO Nº\s*", "", relevant_title)
        logging.info(f"DESPACHO Nº: {relevant_title}")
        check_for_stop_and_pause(stop_event, pause_event)

        portaria_pattern = r"Portaria\s+n(?:\.|[º°])?\s*(\d+,\s*de\s*\d{1,2}\s*(?:de\s*)?([A-Za-zçãÁÉÍÓÚÂÊÎÔÛÀüÜ]+)\s*de\s*\d{4})"
        portaria_match = re.search(portaria_pattern, document_text, re.IGNORECASE)

        if portaria_match:
            # Extract the full date match (group 1) and the month name (group 2)
            full_match = portaria_match.group(1).strip()
            month_name = portaria_match.group(2).strip()

            month_name_pattern_part = re.escape(month_name) # Escape just in case
            
            # Check if the format is missing "de" before the month name
            if re.search(r"\d{1,2}\s+" + month_name_pattern_part, full_match, re.IGNORECASE) and \
                not re.search(r"\d{1,2}\s+de\s+" + month_name_pattern_part, full_match, re.IGNORECASE):
                # Fix the format by adding the missing "de"
                relevant_title2 = re.sub(r"(\d{1,2})\s+(" + month_name_pattern_part + r")", r"\1 de \2", full_match, flags=re.IGNORECASE)
            else:
                relevant_title2 = full_match
                
            # Clean up extra spaces
            relevant_title2 = re.sub(r'\s+', ' ', relevant_title2).strip()
            logging.info(f"PORTARIA Nº: {relevant_title2}")
        else:
            relevant_title2 = None
            logging.warning("Portaria nº information not found in the document.")
        check_for_stop_and_pause(stop_event, pause_event)

        cpf_pattern = r"CPF n[º°]\s*:?\s*(\d{3}\.\d{3}\.\d{3}\s*[-\.]\s*\d{2})"
        cpf_match = re.search(cpf_pattern, document_text, re.IGNORECASE)
        cpf_number = cpf_match.group(1).strip() if cpf_match else None
        if cpf_number:
            # Remove spaces that might exist between hyphen and last digits
            cpf_number = cpf_number.replace(" ", "")
            # Normalize CPF format to ensure standard XXX.XXX.XXX-XX
            cpf_number = re.sub(r'(\d{3})\.(\d{3})\.(\d{3})[.\-](\d{2})', r'\1.\2.\3-\4', cpf_number)
            # Verify normalized format
            if not re.match(r'^\d{3}\.\d{3}\.\d{3}-\d{2}$', cpf_number):
                logging.warning(f"CPF format could not be normalized: {cpf_number}")
        if cpf_number and not re.match(r'^\d{3}\.\d{3}\.\d{3}-\d{2}$', cpf_number):
            logging.warning(f"Extracted CPF does not match expected format: {cpf_number}")
        elif not cpf_number:
            logging.error("CPF number not found in the document text.")
            return None, relevant_title, relevant_title2, None, None, None
        check_for_stop_and_pause(stop_event, pause_event)
        
        start_index = cpf_match.end() if cpf_match else 0
        end_phrase_1 = "cálculos de proventos (Código SEI nº "
        end_phrase_2 = "cálculos elaborados à planilha (Código SEI nº "
        end_phrase_3 = "cálculos de proventos (" # Prefix of end_phrase_1

        # Find indices
        idx1 = document_text.find(end_phrase_1, start_index)
        idx2 = document_text.find(end_phrase_2, start_index)
        idx3 = document_text.find(end_phrase_3, start_index)

        # Store potential matches with their properties
        matches = []
        if idx1 != -1:
            matches.append({'index': idx1, 'phrase': end_phrase_1, 'type': 1, 'priority': 1}) # Highest priority
        if idx2 != -1:
            matches.append({'index': idx2, 'phrase': end_phrase_2, 'type': 2, 'priority': 1}) # Highest priority
        if idx3 != -1:
            matches.append({'index': idx3, 'phrase': end_phrase_3, 'type': 3, 'priority': 2}) # Lower priority

        if not matches:
            logging.warning("No defined end phrase found after the CPF in the document text.")
            return number_after_despacho, relevant_title, relevant_title2, None, cpf_number, None

        matches.sort(key=lambda x: (x['index'], x['priority']))
        
        best_match = None
        if matches:
            if matches[0]['type'] == 3 and idx1 != -1 and matches[0]['index'] == idx1:
                for m in matches:
                    if m['type'] == 1:
                        best_match = m
                        break
                if not best_match: 
                     best_match = matches[0] 
            else:
                best_match = matches[0] 
        
        if not best_match: 
            logging.warning("Logical error: No best match found despite initial matches.")
            return number_after_despacho, relevant_title, relevant_title2, None, cpf_number, None

        end_index = best_match['index']
        found_phrase = best_match['phrase']
        found_phrase_type = best_match['type']

        if found_phrase_type in [1, 2]: # Phrases with "Código SEI nº "
            # Include the full phrase as before, Apostila.py will handle replacement
            chunk_of_text = document_text[start_index : end_index + len(found_phrase)]
            text_after_chunk_start_index = end_index + len(found_phrase)
        elif found_phrase_type == 3: # Phrase "cálculos de proventos ("
            chunk_of_text = document_text[start_index : end_index + (len(found_phrase) -1) ] # Exclude the '('
            text_after_chunk_start_index = end_index + len(found_phrase) # Number starts after '('
        else: # Should not happen if best_match is always set
            logging.error("Undefined found_phrase_type, cannot define chunk_of_text")
            return number_after_despacho, relevant_title, relevant_title2, None, cpf_number, None

        chunk_of_text = bold_selected_words(chunk_of_text)

        # Extract number_in_chunk
        number_in_chunk = None
        search_text_for_number = document_text[text_after_chunk_start_index:]

        if found_phrase_type in [1, 2]:
            full_search_text = document_text[end_index:]
            number_match = re.search(r"\(Código\s*SEI\s*n[ºo°]\s*(\d+)\s*\)", full_search_text)
            if number_match:
                number_in_chunk = number_match.group(1).strip()
        elif found_phrase_type == 3:
            # search_text_for_number for type 3 starts *after* "cálculos de proventos ("
            # So it starts with "NUMBER)"
            number_match = re.search(r"^\s*(\d+)\s*\)", search_text_for_number)
            if number_match:
                number_in_chunk = number_match.group(1).strip()

        if number_in_chunk:
            pass
        else:
            # ... (error handling for not finding number_in_chunk) ...
            context_around_end_phrase = document_text[max(0, end_index - 20) : text_after_chunk_start_index + 50]
            logging.error(f"Could not extract number_in_chunk using phrase type {found_phrase_type}. Context: '...{context_around_end_phrase}...'. Searched in: '{search_text_for_number[:50]}...'")
            return number_after_despacho, relevant_title, relevant_title2, chunk_of_text, cpf_number, None
                            
        return number_after_despacho, relevant_title, relevant_title2, chunk_of_text, cpf_number, number_in_chunk
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Error locating 'Despacho do Gabinete' in the document tree: {e}")
        failed_processes.add(process_number)
        return None, None, None, None, None, None
                            
def check_for_portaria(driver, process_number, failed_processes):
    """Check for Portaria document"""
    number_after_portaria = None
    try:
        # Step 1: Ensure the document tree is expanded
        if not locate_and_expand_tree(driver):
            raise Exception("Failed to expand the document tree.")
        # Step 2: Switch to the document tree iframe
        driver.switch_to.default_content()
        tree_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrArvore"]'))
        )
        driver.switch_to.frame(tree_iframe)
        time.sleep(1)
        # Step 3: Search for "Portaria - GOIASPREV" in the document tree
        document_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]'))
        )
        # Reverse the list to search from bottom to top
        for element in reversed(document_elements):
            element_text = element.text
            if "Portaria - GOIASPREV" in element_text:
                # Extract and log the number after "Portaria - GOIASPREV"
                match = re.search(r'\((\d+)\)', element_text)
                if match:
                    number_after_portaria = match.group(1)
                    logging.info(f"PORTARIA - SEI: {number_after_portaria}")
                else:
                    logging.warning("No number found in 'Portaria - GOIASPREV'.")
                break
        else:
            logging.error("'Portaria - GOIASPREV' not found in the document tree.")
            failed_processes.add(process_number)
        return number_after_portaria  # Return as None if not found
    except Exception as ex:
        logging.error(f"Unexpected error during 'Portaria' processing: {ex}")
        failed_processes.add(process_number)
        save_failed_process(process_number)
    # Return the extracted value
    return number_after_portaria

def check_diario_date(driver, process_number):
    """
    Downloads the Diário Oficial to a temporary location, extracts the date, and cleans up.
    """

    diario_date = None
    temp_download_dir = tempfile.mkdtemp()

    try:
        # Step 1: Ensure the document tree is expanded
        if not locate_and_expand_tree(driver):
            raise Exception("Failed to expand the document tree.")
            
        # Step 2: Switch to the document tree iframe
        driver.switch_to.default_content()
        tree_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrArvore"]'))
        )
        driver.switch_to.frame(tree_iframe)
        
        # Step 3: Search for "Diário Oficial" in the document tree
        document_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]'))
        )
        diario_found = False
        for element in reversed(document_elements):
            if element.text.startswith("Diário Oficial"):
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(1)
                try:
                    element.click()
                except Exception as e:
                    logging.warning(f"Regular click failed: {str(e)}. Trying JavaScript click.")
                    driver.execute_script("arguments[0].click();", element)
                
                diario_found = True
                break
        
        if not diario_found:
            raise Exception("'Diário Oficial' document not found in tree.")

        # --- Configure download to the temporary directory ---
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": temp_download_dir  # Use our temp directory
        })
        
        # Click "open in new tab" to trigger the download
        driver.switch_to.default_content()
        parent_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrConteudoVisualizacao"]'))
        )
        driver.switch_to.frame(parent_iframe)
        document_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrVisualizacao"]'))
        )
        driver.switch_to.frame(document_iframe)
        open_in_new_tab_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="divArvoreInformacao"]/a'))
        )
        open_in_new_tab_button.click()
        
        # --- Wait for the download to complete in the temp folder ---
        downloaded_file_path = None
        wait_time = 0
        max_wait_time = 20 # Wait up to 20 seconds for the download
        while not downloaded_file_path and wait_time < max_wait_time:
            time.sleep(1)
            wait_time += 1
            # Look for any .pdf file that isn't a chrome temp file
            temp_files = [f for f in os.listdir(temp_download_dir) if f.endswith('.pdf') and not f.endswith('.crdownload')]
            if temp_files:
                # Assume the first one is our file
                downloaded_file_path = os.path.join(temp_download_dir, temp_files[0])
        
        if not downloaded_file_path:
            raise Exception("Download of 'Diário Oficial' timed out or failed.")

        # Now you can extract the date or perform any other actions on the temp file
        diario_date = extract_diario_date(downloaded_file_path)
        
    except Exception as e:
        logging.error(f"Error downloading 'Diário Oficial' for process {process_number}: {str(e)}")
        diario_date = None # Ensure it returns None on failure
    finally:
        # --- Cleanup ---
        try:
            shutil.rmtree(temp_download_dir)
        except Exception as e:
            logging.error(f"Failed to remove temporary directory {temp_download_dir}. Error: {e}")
            
    return diario_date

def extract_diario_date(pdf_path):
    """Extract date from Diário Oficial PDF"""
    diario_date = None
    try:
        # Step 1: Open the PDF and extract text from the first page
        with fitz.open(pdf_path) as pdf:
            first_page_text = pdf[0].get_text()  # Extracts text from the first page
        # Step 2: Search for the date pattern in the text
        date_pattern = r'\b[A-ZÀ-ÿ]+, \w+-FEIRA, (\d{1,2} DE [A-ZÀ-ÿ]+ DE \d{4})\b'
        match = re.search(date_pattern, first_page_text)
        # Step 3: If the date pattern is found, extract just the date part
        if match:
            date_only = match.group(1)  # Capture only the actual date part (e.g., "02 DE MAIO DE 2024")
            day_part_match = re.match(r"(\d)\s+DE", date_only)
            if day_part_match and len(day_part_match.group(1)) == 1:
                date_only = "0" + date_only
            diario_date = date_only.lower()  # Convert the date to lowercase
            
            logging.info(f"Diário Oficial de: {diario_date}")
        else:
            raise Exception("Diario date not found on the first page of the PDF.")
    except Exception as e:
        logging.error(f"Error extracting Diario date from PDF: {str(e)}")
    return diario_date

def bold_selected_words(text):
    """Bold selected words in the text"""
    words_to_bold = {"VENCIMENTO", "GRATIFICAÇÃO ADICIONAL", "GRATIFICAÇÃO DE INCENTIVO FUNCIONAL"}
    for word in words_to_bold:
        text = re.sub(rf'\b{word}\b', f'**{word}**', text)
    return text

def add_marker_and_save(driver, process_number, failed_processes, successful_processes):
    """Add marker and save the document"""
    try:
        # Click the "Add" button
        add_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btnAdicionar"]'))
        )
        add_button.click()
        time.sleep(1)
        # Click the dropdown menu and select the "RETIFICAÇÃO - APOSTILAMENTO" option
        dropdown_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="selMarcador"]/div/a'))
        )
        dropdown_menu.click()
        time.sleep(1)
        option_to_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[normalize-space()='RETIFICAÇÃO - APOSTILAMENTO']"))
        )
        option_to_select.click()
        time.sleep(1)
        # Click the save button
        save_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="sbmSalvar"]'))
        )
        save_button.click()
        time.sleep(1)
        # Confirm that the marker was added
        WebDriverWait(driver, 30).until(
            EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div/div/form/div[3]/table/tbody/tr[2]/td[2]'), "RETIFICAÇÃO - APOSTILAMENTO")
        )
        logging.info("Marker successfully saved as 'RETIFICAÇÃO - APOSTILAMENTO'.")
        failed_processes.add(process_number)
        save_failed_process(process_number)
    except TimeoutException:
        logging.error("Timeout while adding the marker - it may not have been saved.")
        return False
    except Exception as e:
        logging.error(f"Error occurred while adding the marker: {e}")
        return False
    finally:
        pass

def remove_marker_and_save(driver, process_number):
    """Remove marker and save the document"""
    try:
        # Switch to the default content before accessing frames
        driver.switch_to.default_content()
        time.sleep(1)

        # Step 1: Access the Tree iFrame and Click on Process Link by process number
        tree_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrArvore"]'))
        )
        driver.switch_to.frame(tree_iframe)
        time.sleep(1)

        # Locate the process link by matching the process number in the span element
        process_number_xpath = f'//span[@class="noVisitado" and text()="{process_number}"]'
        process_number_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, process_number_xpath))
        )
        time.sleep(1)

        # Scroll the element into view
        driver.execute_script("arguments[0].scrollIntoView(true);", process_number_element)
        time.sleep(1)

        # Click on the process link (span element containing the process number)
        process_number_element.click()
        time.sleep(1)

        # Step 2: Switch to Parent iFrame
        driver.switch_to.default_content()
        parent_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrConteudoVisualizacao"]'))
        )
        driver.switch_to.frame(parent_iframe)
        time.sleep(1)

        # Step 3: Click on the Marker Icon by its src attribute
        marker_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//img[contains(@src, "marcador_gerenciar.svg")]'))
        )
        marker_icon.click()
        time.sleep(1)

        document_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrVisualizacao"]'))
        )
        driver.switch_to.frame(document_iframe)
        time.sleep(1)

        # Step 4: Locate and Click the White Marker Checkbox
        white_marker_checkbox = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[1]/div'))
        )
        white_marker_checkbox.click()
        time.sleep(1)

        # Step 5: Click the "Remove" Button
        remove_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btnRemover"]'))
        )
        remove_button.click()
        time.sleep(1)

        # Step 6: Handle the Alert
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        time.sleep(1)
        logging.info("White marker successfully removed and document saved.")
    except Exception as e:
        logging.error(f"An error occurred in remove_marker_and_save: {e}")

def determine_year_range(year):
    """Determine the year range based on the given year"""
    if 1988 <= year <= 1992:
        return "1988"
    elif 1993 <= year <= 1998:
        return "1993"
    elif 1999 <= year <= 2003:
        return "1999"
    elif year == 2004:
        return "2004"
    elif year == 2005:
        return "2005"
    elif 2006 <= year <= 2007:
        return "2006"
    elif 2008 <= year <= 2009:
        return "2008"
    elif year == 2010:
        return "2010"
    return None
