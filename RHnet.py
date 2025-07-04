import os
import re
import time
import logging
import base64
import tempfile
import shutil

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

from utils import start_new_driver_session

# Constants
URL_RHNET = "https://aplicacoes.expresso.go.gov.br/"

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def login_to_rhnet(driver, username, password):
    """Log in to the RHnet system"""

    driver.get(URL_RHNET)
    try:
        login_box = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="usernameUserInput"]')))
        login_box.send_keys(username)
        password_box = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
        password_box.send_keys(password)
        login_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginForm"]/button')))
        login_button.click()
    except Exception as e:
        logging.error(f"Login failed: {e}")
        return False
    try:
        continuar_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="approve"]')))
        continuar_button.click()
    except TimeoutException:
        pass
    return True

def navigate_to_consultar_ficha_financeira(driver):
    """Navigate to the 'Consultar Ficha Financeira' page"""
    try:
        people_icon_xpath = "//i[@class='icone-grid pi pi-users']"
        people_icon = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, people_icon_xpath)))
        people_icon.click()
        time.sleep(2)
        driver.switch_to.frame("menu")
        processamento_button = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[3]')))
        action = ActionChains(driver)
        action.move_to_element(processamento_button).perform()
        processamento_button.click()
        driver.switch_to.default_content()
        driver.switch_to.frame("principal")
        consultar_button = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//div[contains(text(), "Consultar Ficha Financeira")]')))
        action.move_to_element(consultar_button).perform()
        consultar_button.click()
        servidor_button = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="Servidor"]')))
        action.move_to_element(servidor_button).perform()
        servidor_button.click()
    except Exception as e:
        logging.error(f"Navigation failed: {e}")
        return False
    return True

def fill_form_and_select_option(driver, cpf_number, option_index=1):
    """Fill the form and select an option from the dropdown"""
    try:
        # Locate and fill the 'Órgão' textbox
        orgao_textbox = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/form/center[1]/table/tbody/tr[1]/td[2]/input[2]'))
        )
        orgao_textbox.clear()  # Clear any existing value
        orgao_textbox.send_keys("309")
    except Exception as e:
        logging.error(f"Órgão textbox not found or visible: {e}")
        return False, option_index
    try:
        # Locate and fill the 'CPF' textbox
        cpf_textbox = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/input[1]'))
        )
        cpf_textbox.clear()  # Clear any existing value
        cpf_textbox.send_keys(cpf_number)
        time.sleep(1)  # Allow page interactions
    except Exception as e:
        logging.error(f"CPF textbox error: {e}")
        return False, option_index
    while option_index is not None:
        try:
            # Locate the dropdown menu
            dropdown_menu = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/form/center[1]/table/tbody/tr[3]/td[2]/select'))
            )
            select = Select(dropdown_menu)
            options = select.options
            # Check if the desired option exists
            if option_index < len(options):
                option = options[option_index]
                try:
                    # Re-locate the dropdown menu and select the option
                    dropdown_menu = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/form/center[1]/table/tbody/tr[3]/td[2]/select'))
                    )
                    select = Select(dropdown_menu)
                    # Use a more robust method to select the option
                    select.select_by_index(option_index)
                    time.sleep(1)  # Short wait
                    # Press Enter
                    ActionChains(driver).send_keys(Keys.ENTER).perform()
                    time.sleep(2)  # Allow for interactions
                    # Check if CPF field is empty
                    cpf_textbox = driver.find_element(By.XPATH, '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/input[1]')
                    if not cpf_textbox.get_attribute('value').strip():
                        # If we were on "Ativado", switch to "Desativado"
                        if option_index == 1:
                            return False, 2  # Retry with "Desativado"
                        # If we've already tried both options, return failure
                        return False, None
                    # Option selection succeeded
                    return True, option_index
                except StaleElementReferenceException:
                    logging.warning("Stale element detected. Retrying...")
                    time.sleep(1)  # Wait a moment before retrying
                    continue
                except Exception as e:
                    logging.error(f"Error selecting option at index {option_index}: {e}")
                    return False, option_index + 1
            # No valid options left
            logging.warning("No valid options left in the dropdown menu.")
            return False, None
        except Exception as e:
            logging.error(f"Error during dropdown handling: {e}")
            return False, None

def extract_person_info(driver):
    """Extract the person's name from the field to the right of the CPF textbox"""
    try:
        person_name_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/input[2]'))
        )
        person_name = person_name_element.get_attribute('value').strip().upper()
        return person_name
    except Exception as e:
        logging.error(f"Person's name not found in the field to the right of the CPF textbox: {e}")
        return None

def extract_vinculo_year_cargo(driver):
    """Extract vinculo number, year, and cargo from the selected option text"""
    try:
        second_dropdown_menu = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/form/center[1]/table/tbody/tr[4]/td[2]/select'))
        )
        select_second = Select(second_dropdown_menu)
        # Get all options
        options = select_second.options
        # Select the last option (subtract 1 because indices start at 0)
        last_index = len(options) - 1
        select_second.select_by_index(last_index)
        time.sleep(2)  # Allow the selection to load
        # Retrieve the selected option text to find "vinculo_number", "year", and "cargo"
        selected_option_text = select_second.first_selected_option.text
        # Use regex to find the vinculo number within brackets
        vinculo_match = re.search(r'\[(\d+)\]', selected_option_text)
        if vinculo_match:
            vinculo_number = vinculo_match.group(1)
        else:
            logging.error("Vinculo number not found in the selected option text.")
            vinculo_number = None
        # Use regex to find the year in the date at the beginning of the line
        year_match = re.search(r'(\d{2}/\d{2}/(\d{4}))', selected_option_text)
        if year_match:
            year = int(year_match.group(2))
        else:
            logging.error("Year not found in the selected option text.")
            year = None
        # Use regex to find the cargo
        cargo_match = re.search(r'\d{2}/\d{2}/\d{4} - (.*?(Administrativo|Professor|Analista)[\s\S]*?)\s*\[', selected_option_text)
        if cargo_match:
            cargo = cargo_match.group(1).strip()
        else:
            logging.error("Cargo not found in the selected option text.")
            cargo = None
        return vinculo_number, year, cargo
    except Exception as e:
        logging.error(f"Error in automate_RHnet: {e}")
        return None, None, None

def click_consultar_button(driver):
    """Click the 'Consultar' button"""
    try:
        # Wait until the "Consultar" button is clickable
        consultar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/center[2]/input[1]')))
        consultar_button.click()
    except TimeoutException:
        logging.warning("Consultar button not found or not clickable")
    except Exception as e:
        logging.error(f"Error clicking 'Consultar' button: {e}")

def click_checkboxes(driver):
    """Click the checkboxes for the document pages"""
    checkbox_names = ["selReg1", "selReg2", "selReg3"]
    for name in checkbox_names:
        try:
            checkbox = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.NAME, name)))
            checkbox.click()
        except Exception as e:
            logging.error(f"Checkbox '{name}' click failed: {e}")

def click_detalhar_button(driver):
    """Click the 'Detalhar' button"""
    try:
        detalhar_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/center[3]/input[2]')))
        detalhar_button.click()
        time.sleep(2)  # Wait for the page to load
    except Exception as e:
        logging.error(f"Error clicking 'Detalhar' button: {e}")

def save_document_pages(driver, download_dir):
    """Saves each document page as a PDF to the specified directory."""

    for page_number in range(1, 4):
        try:
            result = driver.execute_cdp_cmd('Page.printToPDF', {})
            pdf_data = base64.b64decode(result['data'])
            
            file_name = f'ficha_financeira_{page_number}.pdf'
            file_path = os.path.join(download_dir, file_name)
            
            with open(file_path, 'wb') as f:
                f.write(pdf_data)
            
        except Exception as e:
            logging.error(f"Failed to save page {page_number} using CDP: {e}")
            return False 

        if page_number < 3:
            try:
                recuar_button_xpath = '/html/body/form/center[3]/input[1]'
                recuar_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, recuar_button_xpath))
                )
                old_element_reference = driver.find_element(By.TAG_NAME, 'html')
                recuar_button.click()

                WebDriverWait(driver, 10).until(EC.staleness_of(old_element_reference))
                time.sleep(1)
            
            except Exception as e:
                logging.error(f"Failed to click 'Recuar' button to navigate to page {page_number + 1}: {e}")
                return False
    
    return True 
            
def automate_RHnet(cpf_number, username, password):
    """Automates RHnet, downloads files to a temp dir, and returns its path."""

    rhnet_driver = None
    temp_dir_path = None
    
    try:
        temp_dir_path = tempfile.mkdtemp(prefix="ficha_financeira_")
        rhnet_driver = start_new_driver_session()
        
        if not login_to_rhnet(rhnet_driver, username, password):
            logging.error("Login to RHnet failed.")
            return None, None, None, None, None
        
        if not navigate_to_consultar_ficha_financeira(rhnet_driver):
            logging.error("Navigation to 'Consultar Ficha Financeira' failed.")
            return None, None, None, None, None
            
        success, next_index = fill_form_and_select_option(rhnet_driver, cpf_number)
        if not success:
            logging.error("Form filling and option selection failed.")
            success, next_index = fill_form_and_select_option(rhnet_driver, cpf_number, option_index=next_index)
            if not success:
                logging.error("Failed to select a valid option after retry.")
                return None, None, None, None, None

        person_name = extract_person_info(rhnet_driver)
        if not person_name:
            logging.error("Failed to extract person's name.")
            return None, None, None, None, None
            
        vinculo_number, year, cargo = extract_vinculo_year_cargo(rhnet_driver)
        if not vinculo_number or not year or not cargo:
            logging.error("Failed to extract vinculo number, year, or cargo.")
            return None, None, None, None, None
        
        click_consultar_button(rhnet_driver)
        click_checkboxes(rhnet_driver)
        click_detalhar_button(rhnet_driver)
        
        if not save_document_pages(rhnet_driver, temp_dir_path):
            logging.error("Failed to save Ficha Financeira pages.")
            return None, None, None, None, None

    except Exception as e:
        logging.error(f"An unexpected error occurred during RHnet automation: {e}")
        if temp_dir_path and os.path.exists(temp_dir_path):
            shutil.rmtree(temp_dir_path)
            logging.info(f"Cleaned up temp directory {temp_dir_path} after unexpected error.")
        return None, None, None, None, None
    finally:
        if rhnet_driver:
            rhnet_driver.quit()
            
    return person_name, vinculo_number, year, cargo, temp_dir_path

##################### TEST CODE #####################

if __name__ == "__main__":
    
    cpf_number = "123.456.789-00"  # Replace with the actual CPF number

    test_user = "YOUR_TEST_USER"
    test_pass = "YOUR_TEST_PASS"
    automate_RHnet(None, cpf_number, test_user, test_pass)