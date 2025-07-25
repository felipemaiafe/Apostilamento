import logging
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Constants
MAX_RETRIES = 3
RETRY_DELAY = 2
TEXT_AREA_XPATH = '//*[@id="txaEditor_2357"]/p[2]'

def automate_Apostila(driver, relevant_title2, number_after_portaria, process_number, 
                       person_name, cpf_number, chunk_of_text, relevant_title, 
                       number_after_despacho, vinculo_number, diario_date, number_in_chunk,
                       callbacks):
    """Automates Apostila document creation and verification with retry logic"""
    
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

    def create_apostila_document():
        """Create new Apostila document with retries"""
        for attempt in range(MAX_RETRIES):
            try:
                switch_to_ConteudoVisualizacao_frame()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="divArvoreAcoes"]/a[1]/img'))).click()
                switch_to_visualization_frame()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="tblSeries"]/tbody/tr[3]/td/a[2]'))).click()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/form[1]/div[5]/fieldset/div[1]/div'))).click()
                protocol_field = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="txtProtocoloDocumentoTextoBase"]')))
                protocol_field.clear()
                protocol_field.send_keys("57662222")
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="divOptPublico"]/div/label'))).click()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="btnSalvar"]'))).click()
                logging.info("APOSTILA created successfully")
                return True
            except Exception as e:
                logging.warning(f"Attempt {attempt+1} failed to create Apostila: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error("Failed to create Apostila after maximum retries")
                    return False
                time.sleep(RETRY_DELAY)
        return False

    def insert_formatted_text():
        """Insert formatted text with links and bold formatting"""
        original_window = driver.current_window_handle # Store original window
        try:
            # Wait for editor window and switch to it
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            editor_window = [w for w in driver.window_handles if w != original_window][0]
            driver.switch_to.window(editor_window)
            time.sleep(2.5) 

            # Construct Text
            if "(Código SEI nº " in chunk_of_text: # True for type 1 and 2
                base_chunk = chunk_of_text.split("(Código SEI nº ")[0]
                chunk_of_text_with_number = f"{base_chunk}(Código SEI nº {number_in_chunk})"
            elif chunk_of_text.endswith("cálculos de proventos "): # True for type 3
                chunk_of_text_with_number = f"{chunk_of_text.rstrip()} (Código SEI nº {number_in_chunk})"
            else:
                logging.warning(f"Unexpected chunk_of_text format: '{chunk_of_text}'. Using default construction.")
                chunk_of_text_with_number = f"{chunk_of_text} (Código SEI nº {number_in_chunk})"

            replacement_text = (
                f"O Superintendente de Gestão e Desenvolvimento de Pessoas, da Secretaria de Estado da Educação, "
                f"no uso das atribuições que lhe confere o Decreto de 23/04/2020, publicado no Diário Oficial de 24/04/2020, "
                f"declara que, por Portaria nº {relevant_title2}, evento SEI {number_after_portaria}, "
                f"publicada no Diário Oficial de {diario_date}, conforme Processo nº {process_number}, "
                f"foi concedida a **{person_name}, CPF nº {cpf_number}**, aposentadoria em seu único vínculo "
                f"({vinculo_number}){chunk_of_text_with_number}, e conforme informações constantes do Despacho nº {relevant_title} "
                f"(Código SEI nº {number_after_despacho})."
            )

            # --- Locate, Triple-Click, Delete ---
            try:
                # Wait for the paragraph element to be present
                document_text_element = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, TEXT_AREA_XPATH))
                )

                # Use ActionChains to triple-click and delete
                actions = ActionChains(driver)
                actions.move_to_element(document_text_element) # Move to the element first
                actions.click(document_text_element) # Single click
                actions.double_click(document_text_element) # Followed by double click = triple click
                actions.send_keys(Keys.DELETE) # Press delete key
                actions.perform() # Execute the sequence

                time.sleep(0.5) # Pause after delete to allow editor to update

            except TimeoutException:
                logging.error(f"Timeout finding target paragraph for replacement: {TEXT_AREA_XPATH}")
                raise Exception("Target paragraph for replacement not found.")
            except Exception as clear_err:
                logging.error(f"Error during triple-click/delete sequence: {clear_err}")
                raise Exception("Failed to clear target paragraph using triple-click/delete.")
            # --- End of Triple-Click, Delete ---

            # --- Insert new text using the ORIGINAL element.send_keys approach ---
            def insert_text_with_link(text, number=None):
                parts = text.split("**")
                for i, part in enumerate(parts):
                    if i % 2 == 0:
                        # Use direct send_keys to the element
                        document_text_element.send_keys(part)
                    else:
                        # Toggle bold using ActionChains
                        bold_actions = ActionChains(driver)
                        bold_actions.key_down(Keys.CONTROL).send_keys("b").key_up(Keys.CONTROL).perform()
                        # Send bolded part using direct send_keys
                        document_text_element.send_keys(part)
                        # Toggle bold off
                        bold_actions.key_down(Keys.CONTROL).send_keys("b").key_up(Keys.CONTROL).perform()

                if number:
                    # Add link using ActionChains
                    link_actions = ActionChains(driver)
                    link_actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys('l').key_up(Keys.CONTROL).key_up(Keys.SHIFT).perform()
                    time.sleep(0.5)
                    # Send keys directly to the window/focused element for the link popup
                    link_input_actions = ActionChains(driver)
                    link_input_actions.send_keys(number).perform()
                    time.sleep(0.5)
                    link_input_actions.send_keys(Keys.ENTER).perform()
                    time.sleep(1)

            # Split and insert text parts
            text_parts = replacement_text.split(f"{number_after_portaria}")
            insert_text_with_link(text_parts[0], number_after_portaria)

            remainder_after_portaria = text_parts[1]
            text_parts = remainder_after_portaria.split(f"{process_number}")
            insert_text_with_link(text_parts[0], process_number)

            remainder_after_process = text_parts[1]
            text_parts = remainder_after_process.split(f"{number_in_chunk}")
            insert_text_with_link(text_parts[0], number_in_chunk)

            remainder_after_chunk = text_parts[1]
            text_parts = remainder_after_chunk.split(f"{number_after_despacho}")
            insert_text_with_link(text_parts[0], number_after_despacho)

            insert_text_with_link(text_parts[1]) # Final part


            # Save and close editor
            actions_save = ActionChains(driver)
            actions_save.key_down(Keys.CONTROL).key_down(Keys.ALT).send_keys('s').key_up(Keys.ALT).key_up(Keys.CONTROL).perform()
            time.sleep(4)
            driver.close()
            driver.switch_to.window(original_window) # Switch back to original window
            return True

        except Exception as e:
            logging.error(f"Failed to edit Apostila: {str(e)}")
            # Cleanup: Try to close editor window if it's still open and switch back
            try:
                if driver.current_window_handle != original_window:
                    logging.warning("Attempting cleanup: closing editor window after error.")
                    driver.close()
                    driver.switch_to.window(original_window)
            except Exception as cleanup_e:
                logging.error(f"Error during cleanup after Apostila edit failure: {cleanup_e}")
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

    def verify_apostila_content():
        """Verify all required content exists in Apostila"""
        for attempt in range(MAX_RETRIES):
            try:
                click_last_document_in_tree()  # Ensure the document tree refreshes
                
                switch_to_ConteudoVisualizacao_frame()
                switch_to_visualization_frame()

                document_body = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body')))
                content = document_body.text.lower()

                required_content = [
                    relevant_title2.lower(),
                    str(number_after_portaria).lower(),
                    str(process_number).lower(),
                    person_name.lower(),
                    cpf_number.lower(),
                    chunk_of_text.replace("**", "").lower(),
                    relevant_title.lower(),
                    str(number_after_despacho).lower(),
                    vinculo_number.lower(),
                    diario_date.lower(),
                    str(number_in_chunk).lower()
                ]
                missing = [item for item in required_content if item not in content]

                if not missing:
                    logging.info("Apostila content verification successful")
                    return True # Content is correct, exit function successfully
                
                # If content is missing:
                logging.warning(f"Missing content in Apostila: {missing}")

                if attempt == MAX_RETRIES - 1:
                    # This was the last attempt, log final failure and exit
                    logging.error("Apostila content verification failed after maximum retries (content still missing).")
                    return False
                
                # Re-edit the document if content is missing
                logging.info(f"Attempt {attempt + 1}: Content missing, attempting re-edit...")
                try:
                    # Go back to the frame containing the edit button
                    switch_to_ConteudoVisualizacao_frame()
                    edit_button_xpath = "//img[contains(@src, 'documento_editar_conteudo.svg')]"
                    logging.debug("Locating edit button...")
                    edit_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, edit_button_xpath))
                    )
                    edit_button.click()
                    logging.info("Clicked edit button.")
                    time.sleep(2)

                    # Call the text insertion function
                    logging.info("Calling insert_formatted_text for re-edit...")
                    if not insert_formatted_text():
                        logging.error("Re-edit failed because insert_formatted_text returned False.")
                        return False # If re-edit fails catastrophically, stop verification attempts

                    logging.info("Re-edit attempt finished. Loop will continue to next verification attempt.")

                except Exception as edit_err:
                    logging.error(f"Error during re-edit process on attempt {attempt + 1}: {edit_err}")
                    return False # Exit verification if re-edit step fails

            except Exception as e:
                logging.error(f"Verification attempt {attempt + 1} failed with exception: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error("Verification failed on final attempt due to exception.")
                    return False # Failed on last attempt due to exception
                time.sleep(RETRY_DELAY) # Wait before the next attempt in the loop

        # This line is reached only if the loop finishes without returning True (e.g., all attempts failed)
        logging.error("Exited verify_apostila_content loop without successful verification.")
        return False

    def add_to_signing_block():
        """Add document to signing block with retries"""
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

                dropdown = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
                )
                Select(dropdown).select_by_value("1703955")
                time.sleep(2)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, include_button_xpath))
                ).click()
                time.sleep(2)

                logging.info("Apostila added to signing block successfully") 
                return True

            except Exception as e:
                logging.warning(f"Attempt {attempt+1} failed to add Apostila to signing block: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    logging.error("Failed to add Apostila to signing block after maximum retries")
                    return False # Failed after all retries
                time.sleep(RETRY_DELAY) # Wait before retrying

        # Should only be reached if all retries fail
        logging.error("Exited add_to_signing_block loop without success.")
        return False

    try:
        # Main execution flow
        if not create_apostila_document():
            raise Exception(f"Failed to create Apostila document for process {process_number}")
            
        if not insert_formatted_text():
            raise Exception(f"Failed to insert formatted text for process {process_number}")
            
        if not verify_apostila_content():
            raise Exception(f"Failed to verify Apostila content for process {process_number}")
            
        if not add_to_signing_block():
            raise Exception(f"Failed to add Apostila to signing block for process {process_number}")

        # Final verification in document tree
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="ifrArvore"]')))
        document_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//a[contains(@class, "infraArvoreNo")]')))
        
        apostila_found_in_tree = False
        for element in reversed(document_elements):
            if "Apostila" in element.text:
                # Scroll and click logic...
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(1)
                try:
                    element.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", element)
                
                logging.info("Apostila verified in document tree")
                apostila_found_in_tree = True
                break # Exit the loop once found
        
        # --- PLACEMENT OF SUCCESS REPORTING ---
        if apostila_found_in_tree:
            # If the final verification passed, report success and return True
            callbacks['update_checklist']('Apostila', True)
            return True
        else:
            # If the final verification FAILED, raise an exception to be caught
            raise Exception("Apostila not found in document tree after creation.")
                
    except Exception as e:
        logging.error(f"Critical error in Apostila automation: {str(e)}")
        callbacks['update_checklist']('Apostila', False)
        return False