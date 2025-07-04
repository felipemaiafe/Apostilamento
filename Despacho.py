import logging
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# Constants
MAX_RETRIES = 3
RETRY_DELAY = 2
TEXT_AREA_XPATH = '//*[@id="txaEditor_474"]/p/strong'

def automate_Despacho(driver, cpf_number, process_number, callbacks):
    """Automates Despacho document creation and verification with retry logic"""

    def switch_to_ConteudoVisualizacao_frame():
        """Switch to main visualization frame"""
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="ifrConteudoVisualizacao"]')))
        time.sleep(0.5)

    def switch_to_visualization_frame():
        """Switch to main visualization frame"""
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="ifrVisualizacao"]')))
        time.sleep(0.5)

    def create_despacho_document():
        """Create new Despacho document with retries"""
        for attempt in range(MAX_RETRIES):
            try:
                switch_to_ConteudoVisualizacao_frame()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="divArvoreAcoes"]/a[1]/img'))).click()
                switch_to_visualization_frame()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="tblSeries"]/tbody/tr[14]/td/a[2]'))).click()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/form[1]/div[5]/fieldset/div[1]/div'))).click()
                protocol_field = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="txtProtocoloDocumentoTextoBase"]')))
                protocol_field.clear()
                protocol_field.send_keys("57689116")
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="divOptPublico"]/div/label'))).click()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="btnSalvar"]'))).click()
                logging.info("DESPACHO created successfully")
                return True
            except Exception as e:
                logging.warning(f"Attempt {attempt+1} failed to create Despacho: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error("Failed to create Despacho after maximum retries")
                    return False
                time.sleep(RETRY_DELAY)
        return False

    def update_cpf_number():
        """Update CPF number in document with retries"""
        original_window = driver.current_window_handle # Store original window
        try:
            # Wait for editor window and switch to it
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            editor_window = [w for w in driver.window_handles if w != original_window][0]
            driver.switch_to.window(editor_window)
            time.sleep(2.5) 

            cpf_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, TEXT_AREA_XPATH)))
            
            # Clear existing CPF and insert new one
            actions = ActionChains(driver)
            actions.move_to_element(cpf_element).click().perform()
            time.sleep(0.5)

            actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
            time.sleep(0.5)

            actions.send_keys(Keys.DELETE).perform()
            time.sleep(0.5)

            actions.send_keys(f"CPF: {cpf_number}").perform()
            time.sleep(0.5)

            # Save and close editor
            actions_save = ActionChains(driver)
            actions_save.key_down(Keys.CONTROL).key_down(Keys.ALT).send_keys('s').key_up(Keys.ALT).key_up(Keys.CONTROL).perform()
            time.sleep(4)
            driver.close()
            driver.switch_to.window(original_window) # Switch back to original window
            return True        

        except Exception as e:
            logging.error(f"Failed to edit Despacho: {str(e)}")
            # Cleanup: Try to close editor window if it's still open and switch back
            try:
                if driver.current_window_handle != original_window:
                    logging.warning("Attempting cleanup: closing editor window after error.")
                    driver.close()
                    driver.switch_to.window(original_window)
            except Exception as cleanup_e:
                logging.error(f"Error during cleanup after Despacho edit failure: {cleanup_e}")
            return False
        
    def click_last_document_in_tree():
        """Clicks on the last document in the document tree to force a refresh."""
        try:
            driver.switch_to.default_content()
            tree_iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ifrArvore"]')))
            driver.switch_to.frame(tree_iframe)
            
            document_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]')))
            
            if document_elements:
                # Scroll to the element before clicking
                driver.execute_script("arguments[0].scrollIntoView(true);", document_elements[-1])
                time.sleep(1)
                
                try:
                    document_elements[-1].click()
                except Exception as e:
                    logging.warning(f"Regular click failed: {str(e)}. Trying JavaScript click.")
                    driver.execute_script("arguments[0].click();", document_elements[-1])
                
                time.sleep(2)  # Give time for the refresh to take effect
        except Exception as e:
            logging.error(f"Failed to click the last document in the tree: {str(e)}")

    def verify_despacho_content():
        """Verify CPF number in document content"""
        for attempt in range(MAX_RETRIES):
            try:
                click_last_document_in_tree()  # Ensure the document tree refreshes

                switch_to_ConteudoVisualizacao_frame()
                switch_to_visualization_frame()

                document_body = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body'))
                )
                content = document_body.text
                expected_text = f"CPF: {cpf_number}"

                if expected_text in content:
                    logging.info("Despacho CPF verification successful")
                    return True # Content is correct, exit successfully

                # If content is missing:
                logging.warning(f"Attempt {attempt + 1}: Expected text '{expected_text}' not found in Despacho content.")

                if attempt == MAX_RETRIES - 1:
                    # This was the last attempt, log final failure and exit
                    logging.error("Despacho verification failed after maximum retries (content still missing).")
                    return False

                # --- Re-edit logic  ---
                logging.info(f"Attempt {attempt + 1}: Content missing, attempting re-edit...")
                try:
                    # Go back to the frame containing the edit button
                    switch_to_ConteudoVisualizacao_frame()
                    edit_button_xpath = "//img[contains(@src, 'documento_editar_conteudo.svg')]"
                    logging.debug("Locating edit button for Despacho...")
                    edit_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, edit_button_xpath))
                    )
                    edit_button.click()
                    logging.info("Clicked edit button for Despacho.")
                    time.sleep(2)

                    # Call the CPF update function
                    logging.info("Calling update_cpf_number for re-edit...")
                    if not update_cpf_number():
                        logging.error("Re-edit failed because update_cpf_number returned False.")
                        return False # If re-edit fails catastrophically, stop verification attempts

                    logging.info("Re-edit attempt finished. Loop will continue to next verification attempt.")

                except Exception as edit_err:
                    logging.error(f"Error during Despacho re-edit process on attempt {attempt + 1}: {edit_err}")
                    # If re-edit fails due to an exception, stop verification
                    return False

            except Exception as e:
                logging.error(f"Despacho Verification attempt {attempt + 1} failed with exception: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error("Despacho Verification failed on final attempt due to exception.")
                    return False # Failed on last attempt due to exception
                time.sleep(RETRY_DELAY) # Wait before the next attempt in the loop

        # This line is reached only if the loop finishes without returning True
        logging.error("Exited verify_despacho_content loop without successful verification.")
        return False

    def add_to_signing_blocks():
        """Add document to signing blocks."""
        for attempt in range(MAX_RETRIES):
            try:
                switch_to_ConteudoVisualizacao_frame()

                add_to_block_icon_xpath = "//img[contains(@src, 'bloco_incluir_protocolo.svg')]"

                add_button_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, add_to_block_icon_xpath))
                )
                add_button_element.click()

                switch_to_visualization_frame()

                # Wait for dropdown and select value
                dropdown_xpath = '//*[@id="selBloco"]'
                include_button_xpath = '//*[@id="sbmIncluir"]'
                
                # Select block 1468569 from the dropdown list
                dropdown = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
                )
                Select(dropdown).select_by_value("1468569")
                time.sleep(2)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, include_button_xpath))
                ).click()
                time.sleep(2) 

                # Select block 1334482 from the dropdown list
                dropdown = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
                )
                Select(dropdown).select_by_value("1334482")
                time.sleep(2)

                # Click the last checkbox to mark the document
                checkboxes = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, '//input[starts-with(@id, "chkDocumentosItem")]')))
                if checkboxes:
                    last_checkbox = max(checkboxes, key=lambda x: int(x.get_attribute('id').split('Item')[-1]))
                    driver.execute_script("arguments[0].click();", last_checkbox)

                # Click "Incluir" button again
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, include_button_xpath))
                ).click()
                time.sleep(2)

                logging.info("Despacho added to both signing blocks successfully")
                return True

            except Exception as e:
                logging.warning(f"Attempt {attempt+1} failed to add Despacho to signing block: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error("Failed to add Despacho to signing block after maximum retries")
                    return False # Failed after all retries
                time.sleep(RETRY_DELAY) # Wait before retrying

        # Should only be reached if all retries fail
        logging.error("Exited add_to_signing_block loop without success.")
        return False
        
    try:
        # Main execution flow
        if not create_despacho_document():
            raise Exception(f"Failed to create Despacho for process {process_number}")

        if not update_cpf_number() or not verify_despacho_content():
            raise Exception(f"Failed to update or verify CPF in Despacho for process {process_number}")

        if not add_to_signing_blocks():
            raise Exception(f"Failed to add Despacho to signing blocks for process {process_number}")

        # Final verification in document tree
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="ifrArvore"]')))
        document_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]')))
        
        despacho_found_in_tree = False
        for element in reversed(document_elements):
            if "Despacho" in element.text:
                # Scroll and click logic...
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(1)
                try:
                    element.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", element)
                
                logging.info("Despacho verified in document tree")
                despacho_found_in_tree = True
                break # Exit the loop once found
        
        # --- PLACEMENT OF SUCCESS REPORTING ---
        if despacho_found_in_tree:
            # If the final verification passed, report success and return True
            callbacks['update_checklist']('Despacho', True)
            return True
        else:
            # If the final verification FAILED, raise an exception to be caught
            raise Exception("Despacho not found in document tree after creation.")

    except Exception as e:
        logging.error(f"Critical error in Despacho automation: {str(e)}")
        callbacks['update_checklist']('Despacho', False)
        return False