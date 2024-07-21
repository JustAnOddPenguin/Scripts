import os
import time
import logging
import colorlog
from datetime import datetime, timedelta
import xlwings as xw
import win32com.client as win32
import customtkinter as ctk
from tkinter import filedialog, messagebox, Text
import threading

# Function to create the log directory if it doesn't exist
def ensure_log_directory_exists(log_directory):
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)

# Function to format the log file name and ensure log directory exists
def get_log_file_name(log_directory, suffix=''):
    ensure_log_directory_exists(log_directory)
    date_str = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    log_file_name = os.path.join(log_directory, f'{date_str}{suffix}.log')
    return log_file_name

# Configure logging
def configure_logging(directory, suffix='', text_widget=None):
    log_directory = r'C:\temp\FileFlowLogs'
    log_file_name = get_log_file_name(log_directory, suffix)
    log_formatter = colorlog.ColoredFormatter(
        '%(log_color)s%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        reset=True,
        log_colors={
            'DEBUG': 'cyan',
            'INFO': 'green',
            'WARNING': 'yellow',
            'ERROR': 'red',
            'CRITICAL': 'bold_red',
        }
    )

    file_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Remove previous handlers
    if logger.hasHandlers():
        logger.handlers.clear()

    # Console handler with color
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)

    # File handler
    file_handler = logging.FileHandler(log_file_name, mode='w')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    # Text widget handler
    if text_widget:
        text_handler = TextHandler(text_widget, app)  # Pass the main application window to the handler
        text_handler.setFormatter(file_formatter)  # Use file formatter to avoid color codes
        logger.addHandler(text_handler)

class TextHandler(logging.Handler):
    def __init__(self, text_widget, root):
        super().__init__()
        self.text_widget = text_widget
        self.root = root  # Store a reference to the main application window

        # Create tags for different log levels
        self.text_widget.tag_configure("DEBUG", foreground="cyan")
        self.text_widget.tag_configure("INFO", foreground="green")
        self.text_widget.tag_configure("WARNING", foreground="yellow")
        self.text_widget.tag_configure("ERROR", foreground="red")
        self.text_widget.tag_configure("CRITICAL", foreground="red", font=('Helvetica', '12', 'bold'))

    def emit(self, record):
        msg = self.format(record)
        level = record.levelname

        def append_text():
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n\n', level)  # Add blank line after each log entry and apply tag
            self.text_widget.configure(state='disabled')
            self.text_widget.yview('end')
        self.root.after(0, append_text)  # Call after on the main application window

def detect_files(directory, cutoff_date, file_types=['.xls', '.doc']):
    logger = logging.getLogger()
    logger.info(f"Starting detection in directory: {directory}")
    logger.info(f"Cutoff date for file modification: {cutoff_date.strftime('%Y-%m-%d')}")

    total_files_checked = 0
    files_to_check = []

    try:
        for root, dirs, files in os.walk(directory):
            for file in files:
                total_files_checked += 1
                file_path = os.path.join(root, file)
                last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                logger.debug(f"Checking file: {file_path}, Last Modified: {last_modified_date}")

                if any(file.lower().endswith(ft.lower()) for ft in file_types):
                    if last_modified_date > cutoff_date:
                        files_to_check.append(file_path)
                        logger.info(f"Detected file: {file_path}")
                    else:
                        logger.debug(f"File {file_path} skipped, last modified date {last_modified_date} is before cutoff {cutoff_date}")
                else:
                    logger.debug(f"File {file_path} does not match the file types {file_types}")

    except Exception as e:
        logger.error(f"Error during detection: {e}")

    logger.info(f"Detection completed. Total files checked: {total_files_checked}, Files detected: {len(files_to_check)}")
    return files_to_check

def convert_files(directory, cutoff_date, delay=2, file_types=['.xls', '.doc'], delete_originals=False):
    logger = logging.getLogger()
    logger.info(f"Starting conversion in directory: {directory}")
    logger.info(f"Cutoff date for file modification: {cutoff_date.strftime('%Y-%m-%d')}")

    total_files_checked = 0
    total_files_converted = 0

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    word = win32.Dispatch("Word.Application")
    word.Visible = False

    files_to_convert = detect_files(directory, cutoff_date, file_types)

    try:
        for file_path in files_to_convert:
            total_files_checked += 1
            logger.info(f"Processing file: {file_path}")

            try:
                if file_path.lower().endswith('.xls'):
                    wb = app.books.open(os.path.abspath(file_path))
                    logger.info(f"Opened file: {file_path}")

                    new_file_path = os.path.abspath(file_path.rsplit('.', 1)[0] + '.xlsx')
                    logger.info(f"Converting to: {new_file_path}")

                    wb.save(new_file_path)
                    logger.info(f"Saved file as: {new_file_path}")

                    wb.close()
                    logger.info(f"Closed file: {file_path}")

                elif file_path.lower().endswith('.doc'):
                    doc = word.Documents.Open(os.path.abspath(file_path))
                    logger.info(f"Opened file: {file_path}")

                    new_file_path = os.path.abspath(file_path.rsplit('.', 1)[0] + '.docx')
                    logger.info(f"Converting to: {new_file_path}")

                    doc.SaveAs(new_file_path, FileFormat=16)
                    doc.Close()
                    logger.info(f"Saved file as: {new_file_path}")

                if delete_originals:
                    os.remove(os.path.abspath(file_path))
                    logger.info(f"Deleted original file: {file_path}")

                total_files_converted += 1

            except Exception as e:
                logger.error(f"Error processing file {file_path}: {e}")

            time.sleep(delay)

    finally:
        app.quit()
        word.Quit()

    logger.info(f"Conversion completed. Total files checked: {total_files_checked}, Files converted: {total_files_converted}")

# GUI Application
class Application(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Legacy file checker and converter")
        self.geometry("1024x600")
        self.resizable(True, True)

        # Adjusted the layout to prevent layout issues
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=1)
        self.grid_rowconfigure(6, weight=1)

        self.label_dir = ctk.CTkLabel(self, text="Target Directory:")
        self.label_dir.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_dir = ctk.CTkEntry(self, width=500)
        self.entry_dir.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky="w")
        self.button_browse = ctk.CTkButton(self, text="Browse", command=self.browse_directory)
        self.button_browse.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        self.label_days = ctk.CTkLabel(self, text="Number of Days for Cutoff Date:")
        self.label_days.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_days = ctk.CTkEntry(self, width=100)
        self.entry_days.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_delay = ctk.CTkLabel(self, text="Processing Time per File (seconds):")
        self.label_delay.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_delay = ctk.CTkEntry(self, width=100)
        self.entry_delay.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.label_file_types = ctk.CTkLabel(self, text="File Types:")
        self.label_file_types.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.var_doc = ctk.BooleanVar()
        self.check_doc = ctk.CTkCheckBox(self, text="*.doc", variable=self.var_doc)
        self.check_doc.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.var_xls = ctk.BooleanVar()
        self.check_xls = ctk.CTkCheckBox(self, text="*.xls", variable=self.var_xls)
        self.check_xls.grid(row=3, column=2, padx=5, pady=5, sticky="w")

        self.label_operation = ctk.CTkLabel(self, text="Operation:")
        self.label_operation.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.var_operation = ctk.StringVar(value="check")
        self.radio_check = ctk.CTkRadioButton(self, text="Check", variable=self.var_operation, value="check")
        self.radio_check.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.radio_convert = ctk.CTkRadioButton(self, text="Convert", variable=self.var_operation, value="convert")
        self.radio_convert.grid(row=4, column=2, padx=5, pady=5, sticky="w")

        self.var_delete_originals = ctk.BooleanVar()
        self.check_delete_originals = ctk.CTkCheckBox(self, text="Delete original files after conversion", variable=self.var_delete_originals)
        self.check_delete_originals.grid(row=4, column=3, padx=5, pady=5, sticky="w")

        self.button_run = ctk.CTkButton(self, text="Run", command=self.run_operation)
        self.button_run.grid(row=5, column=0, columnspan=4, padx=10, pady=20)

        self.button_help = ctk.CTkButton(self, text="Help", command=self.show_help)
        self.button_help.grid(row=5, column=1, columnspan=4, padx=10, pady=20)

        self.text_log = Text(self, state='disabled', width=1100, height=300)
        self.text_log.grid(row=6, column=0, columnspan=4, padx=10, pady=5, sticky="nsew")
        self.text_log.configure(bg="black", fg="white")

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_dir.delete(0, ctk.END)
            self.entry_dir.insert(0, directory)

    def run_operation(self):
        directory = self.entry_dir.get().strip()
        if not directory:
            messagebox.showerror("Error", "Please select a target directory.")
            return

        try:
            days = int(self.entry_days.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number of days.")
            return

        try:
            delay = int(self.entry_delay.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid processing time per file.")
            return

        cutoff_date = datetime.now() - timedelta(days=days)
        file_types = []
        if self.var_doc.get():
            file_types.append('.doc')
        if self.var_xls.get():
            file_types.append('.xls')
        if not file_types:
            messagebox.showerror("Error", "Please select at least one file type.")
            return

        operation = self.var_operation.get()
        delete_originals = self.var_delete_originals.get()

        configure_logging(directory, f' - {operation.capitalize()}', self.text_log)

        if operation == 'check':
            threading.Thread(target=self.run_check, args=(directory, cutoff_date, file_types)).start()
        elif operation == 'convert':
            self.run_conversion_thread(directory, cutoff_date, delay, file_types, delete_originals)
        else:
            messagebox.showerror("Error", "Invalid operation. Please select 'check' or 'convert'.")

    def run_check(self, directory, cutoff_date, file_types):
        detect_files(directory, cutoff_date, file_types)
        self.show_completion_message()

    def run_conversion_thread(self, directory, cutoff_date, delay, file_types, delete_originals):
        def conversion_wrapper():
            convert_files(directory, cutoff_date, delay, file_types, delete_originals)
            self.show_completion_message()

        threading.Thread(target=conversion_wrapper).start()

    def show_completion_message(self):
        log_directory = r'C:\temp\FileFlowLogs'
        completion_message = f"Operation completed successfully.\n\nLogs can be found in:\n{log_directory}"
        self.after(0, lambda: messagebox.showinfo("Completed", completion_message))

    def show_help(self):
        help_window = ctk.CTkToplevel(self)
        help_window.title("Help")
        help_window.geometry("650x300")
        help_text = """
        Usage Examples:

        1. Check Files:
        - Select the target directory using the 'Browse' button.
        - Enter the number of days for the cutoff date.
        - Select the file types you want to check (*.doc, *.xls).
        - Choose the 'Check' operation.
        - Click 'Run' to start checking the files.

        2. Convert Files:
        - Select the target directory using the 'Browse' button.
        - Enter the number of days for the cutoff date.
        - Enter the processing time per file in seconds.
        - Select the file types you want to convert (*.doc, *.xls).
        - Choose the 'Convert' operation.
        - Check the 'Delete original files after conversion' if you want to delete the original files after conversion.
        - Click 'Run' to start converting the files.
        """
        ctk.CTkLabel(help_window, text=help_text, justify="left").pack(padx=10, pady=10)

        help_window.lift()  # Bring the window to the front
        help_window.focus_force()  # Give it focus
        help_window.attributes("-topmost", True)  # Keep the window on top

if __name__ == '__main__':
    app = Application()
    app.mainloop()
