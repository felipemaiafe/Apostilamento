import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import logging
import sys
import traceback

# --- GuiLoggingHandler Class ---
class GuiLoggingHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M:%S')

    def emit(self, record):
        msg = self.format(record)
        if self.text_widget.winfo_exists():
            self.text_widget.after(0, self.write, msg + '\n')

    def write(self, msg):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg)
        self.text_widget.configure(state='disabled')
        self.text_widget.see(tk.END)

# --- LoginWindow Class ---
class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login de Acesso")
        self.geometry("320x500")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.credentials = None
        self.create_widgets()
        
        # Center the window on screen
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
        self.focus_force()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        style = ttk.Style(self)
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8)

        sei_frame = ttk.LabelFrame(main_frame, text="Login SEI", padding=15)
        sei_frame.pack(fill="x", pady=(0, 20))
        ttk.Label(sei_frame, text="Usuário:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.sei_user_entry = ttk.Entry(sei_frame, width=40)
        self.sei_user_entry.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(sei_frame, text="Senha:").grid(row=2, column=0, sticky="w", pady=(0, 5))
        self.sei_pass_entry = ttk.Entry(sei_frame, show="*", width=40)
        self.sei_pass_entry.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(sei_frame, text="Órgão:").grid(row=4, column=0, sticky="w", pady=(0, 5))
        sei_orgao_entry = ttk.Entry(sei_frame, width=40)
        sei_orgao_entry.insert(0, "SEDUC")
        sei_orgao_entry.config(state="readonly")
        sei_orgao_entry.grid(row=5, column=0, sticky="ew")

        expresso_frame = ttk.LabelFrame(main_frame, text="Login Goiás Expresso (RHnet)", padding=15)
        expresso_frame.pack(fill="x")
        ttk.Label(expresso_frame, text="Usuário:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.rhnet_user_entry = ttk.Entry(expresso_frame, width=40)
        self.rhnet_user_entry.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(expresso_frame, text="Senha:").grid(row=2, column=0, sticky="w", pady=(0, 5))
        self.rhnet_pass_entry = ttk.Entry(expresso_frame, show="*", width=40)
        self.rhnet_pass_entry.grid(row=3, column=0, sticky="ew")

        submit_button = ttk.Button(main_frame, text="Acessar Automação", command=self.submit, style="TButton")
        submit_button.pack(pady=25, ipady=5)

        self.sei_user_entry.focus_set()
        self.bind("<Return>", lambda event: self.submit())

    def submit(self):
        sei_user = self.sei_user_entry.get().strip()
        sei_pass = self.sei_pass_entry.get().strip()
        rhnet_user = self.rhnet_user_entry.get().strip()
        rhnet_pass = self.rhnet_pass_entry.get().strip()

        if not all([sei_user, sei_pass, rhnet_user, rhnet_pass]):
            messagebox.showerror("Erro de Preenchimento", "Todos os campos de login e senha devem ser preenchidos.", parent=self)
            return

        self.credentials = {
            "sei_user": sei_user,
            "sei_pass": sei_pass,
            "rhnet_user": rhnet_user,
            "rhnet_pass": rhnet_pass,
        }
        self.destroy()

    def on_closing(self):
        self.credentials = None
        self.destroy()


# --- Main Application Class ---
class AutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()  # Keep the main window hidden until it's ready to be shown
        self.title("Automação de Apostilamento SEI")
        self.geometry("800x650") # Set initial size
        self.automation_thread = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.credentials = None
        self.is_running = False
        self.is_paused = False
        self.checklist_vars = {}
        self.processes_analyzed_var = tk.IntVar(value=0)
        self.create_widgets()
        self.configure_logging()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def set_credentials(self, credentials):
        self.credentials = credentials
        
    def show_and_center(self):
        """Calculates position to center the window on screen and shows it."""
        self.update_idletasks() # Ensure window size is calculated
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        self.deiconify() # Make the window visible
    
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)
        main_frame.rowconfigure(2, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        style = ttk.Style(self)
        style.configure("TButton", font=("Segoe UI", 12, "bold"), padding=10)
        
        self.start_stop_button = ttk.Button(button_frame, text="Start", command=self.toggle_automation, style="TButton", width=15)
        self.start_stop_button.pack(side="left", padx=10)
        
        self.pause_resume_button = ttk.Button(button_frame, text="Pause", command=self.toggle_pause, style="TButton", width=15, state="disabled")
        self.pause_resume_button.pack(side="left", padx=10)
        
        checklist_frame = ttk.LabelFrame(main_frame, text="Progresso do Processo Atual", padding="10")
        checklist_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
        
        checklist_items = ["Edital CAPA", "Edital LISTA", "Ficha Financeira", "Apostila", "Despacho"]
        for i, item in enumerate(checklist_items):
            var = tk.StringVar(value="⬜")
            self.checklist_vars[item] = var
            label = ttk.Label(checklist_frame, textvariable=var, font=("Segoe UI", 11))
            label.grid(row=i, column=0, sticky="w")
            text_label = ttk.Label(checklist_frame, text=item, font=("Segoe UI", 11))
            text_label.grid(row=i, column=1, sticky="w", padx=5)
            
        counter_frame = ttk.LabelFrame(main_frame, text="Status Geral", padding="10")
        counter_frame.grid(row=1, column=1, sticky="nsew", padx=(5, 0))
        
        ttk.Label(counter_frame, text="Processos Analisados:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.counter_label = ttk.Label(counter_frame, textvariable=self.processes_analyzed_var, font=("Segoe UI", 24, "bold"))
        self.counter_label.pack(pady=10)
        
        log_frame = ttk.LabelFrame(main_frame, text="Logs", padding="10")
        log_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        
        self.log_widget = scrolledtext.ScrolledText(log_frame, state='disabled', wrap=tk.WORD, font=("Consolas", 9))
        self.log_widget.grid(row=0, column=0, sticky="nsew")

    def configure_logging(self):
        root_logger = logging.getLogger()
        if root_logger.hasHandlers():
            root_logger.handlers.clear()
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', stream=sys.stdout)
        gui_handler = GuiLoggingHandler(self.log_widget)
        root_logger.addHandler(gui_handler)
        root_logger.setLevel(logging.INFO)

    def update_checklist(self, item, success):
        if item in self.checklist_vars:
            current_status = self.checklist_vars[item].get()
            if current_status == "✅":
                return 
            icon = "✅" if success else "❌"
            self.checklist_vars[item].set(icon)

    def reset_checklist(self):
        for var in self.checklist_vars.values():
            var.set("⬜")
        self.update_idletasks()

    def increment_counter(self):
        self.processes_analyzed_var.set(self.processes_analyzed_var.get() + 1)

    def toggle_automation(self):
        if self.is_running:
            self.stop_automation_signal()
        else:
            self.start_automation()

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_event.set() 
            self.pause_resume_button.config(text="Resume")
        else:
            self.pause_event.clear() 
            self.pause_resume_button.config(text="Pause")

    def start_automation(self):
        self.is_running = True
        self.is_paused = False
        self.start_stop_button.config(text="Stop")
        self.pause_resume_button.config(text="Pause", state='normal')
        self.log_widget.configure(state='normal')
        self.log_widget.delete('1.0', tk.END)
        self.log_widget.configure(state='disabled')
        self.processes_analyzed_var.set(0)
        self.reset_checklist()
        self.stop_event.clear()
        self.pause_event.clear()
        self.automation_thread = threading.Thread(
            target=self.run_automation_logic,
            daemon=True
        )
        self.automation_thread.start()

    def stop_automation_signal(self):
        if self.automation_thread and self.automation_thread.is_alive():
            logging.warning("Stop signal sent. Waiting for current task to finish...")
            self.start_stop_button.config(text="Stopping...", state='disabled')
            self.pause_resume_button.config(state='disabled') 
            if self.is_paused: 
                self.pause_event.clear()
            self.stop_event.set()

    def on_automation_finished(self):
        self.is_running = False
        self.is_paused = False
        self.start_stop_button.config(text="Start", state='normal')
        self.pause_resume_button.config(text="Pause", state='disabled') 
        logging.info("Automation process has finished.")

    def run_automation_logic(self):
        try:
            callbacks = {
                'update_checklist': self.update_checklist,
                'reset_checklist': self.reset_checklist,
                'increment_counter': self.increment_counter
            }
            start_loop_modified_for_gui(self.stop_event, self.pause_event, callbacks, self.credentials)

        except Exception as e:
            logging.error(f"Critical error in automation thread: {e}", exc_info=True)
        finally:
            self.after(0, self.on_automation_finished)

    def on_closing(self):
        if self.is_running:
            self.stop_automation_signal()
        self.destroy()

def start_loop_modified_for_gui(stop_event, pause_event, callbacks, credentials):
    logging.info("Starting automation loop.")
    from utils import start_new_driver_session, save_failed_process, load_failed_processes, load_successful_processes
    from Apostilamento import login_to_system, initial_navigate_and_filter, process_navigation, main_workflow, return_to_filtered_list_view, check_for_stop_and_pause
    
    failed_processes = load_failed_processes()
    successful_processes = load_successful_processes()
    driver = None
    try:
        driver = start_new_driver_session()
        if not login_to_system(driver, credentials['sei_user'], credentials['sei_pass']):
            logging.error("Initial login failed.")
            if driver: driver.quit()
            return
        if not initial_navigate_and_filter(driver):
            logging.error("Initial navigation to filtered process list failed.")
            if driver: driver.quit()
            return
        while not stop_event.is_set():
            callbacks['reset_checklist']()
            check_for_stop_and_pause(stop_event, pause_event)
            process_number = None
            try:
                failed_processes.update(load_failed_processes())
                successful_processes.update(load_successful_processes())
                process_number = process_navigation(driver, failed_processes, successful_processes, stop_event, pause_event)
                if process_number is False:
                    logging.info("Automation complete: No more processes found.")
                    break
                elif not process_number:
                    logging.warning("Could not find a suitable process. Will try again.")
                    if not return_to_filtered_list_view(driver): break
                    continue
                logging.info(f"#########################")
                logging.info(f"Processo: {process_number}")
                logging.info(f"#########################")
                main_workflow(
                    driver, process_number, failed_processes, successful_processes,
                    callbacks, credentials, stop_event, pause_event
                )
                callbacks['increment_counter']()
            except Exception as e:
                if type(e).__name__ == 'StopRequestException':
                    logging.info("Stop request confirmed. Exiting main processing loop.")
                    break
                else:
                    logging.error(f"Error during processing loop for process {process_number}: {e}", exc_info=True)
                    if process_number and process_number not in successful_processes:
                        logging.warning(f"Adding process {process_number} to failed list due to exception.")
                        failed_processes.add(process_number)
                        save_failed_process(process_number)
                        callbacks['increment_counter']() 
            finally:
                if stop_event.is_set(): break
                if process_number is False: break
                if not return_to_filtered_list_view(driver): break
    except Exception as outer_e:
        logging.error(f"Critical error in automation logic: {outer_e}", exc_info=True)
    finally:
        if driver:
            driver.quit()
            logging.info("Browser session closed.")
        logging.info("Automation loop has terminated.")

# --- Main entry point ---
if __name__ == "__main__":
    try:
        # 1. Create and show the standalone login window first.
        login_window = LoginWindow()
        # This starts the event loop for the login window and waits until it's closed.
        login_window.mainloop()
        
        # 2. After the login window is closed, check if credentials were provided.
        if login_window.credentials:
            # 3. If yes, create the main application.
            app = AutomationApp()
            app.set_credentials(login_window.credentials)
            
            # Center and show the main window.
            app.show_and_center()

            # Start the main application's event loop.
            app.mainloop()
        # If no credentials were provided, the program simply exits.

    except Exception as e:
        # A final catch-all for any unexpected errors during startup.
        error_message = f"A critical error occurred on startup:\n\n{e}"
        traceback.print_exc()
        try:
            # Try to show a GUI message box.
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Fatal Error", error_message)
            root.destroy()
        except:
            # Fallback to console if GUI fails.
            print("\n" + "="*50)
            print("FATAL ERROR - COULD NOT INITIALIZE GUI")
            print(error_message)
            print("="*50)
            input("\nPress Enter to exit...")