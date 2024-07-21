import os
import time
import logging
import colorlog
from datetime import datetime, timedelta
import xlwings as xw
from docx import Document
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# Function to get the script directory
def get_script_directory():
    return os.path.dirname(os.path.abspath(__file__))

# Function to format the log file name and ensure log directory exists
def get_log_file_name(directory, suffix=''):
    script_dir = get_script_directory()
    log_folder = os.path.join(script_dir, 'logs')
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)
    date_str = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    target_dir_name = os.path.basename(os.path.normpath(directory))
    log_file_name = os.path.join(log_folder, f'{date_str} - {target_dir_name}{suffix}.log')
    return log_file_name

# Configure logging
def configure_logging(directory, suffix='', text_widget=None):
    log_file_name = get_log_file_name(directory, suffix)
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

    def emit(self, record):
        msg = self.format(record)
        def append_text():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n\n')  # Add blank line after each log entry
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
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
                if any(file.lower().endswith(ft.lower()) for ft in file_types):
                    file_path = os.path.join(root, file)
                    last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                    logger.debug(f"File: {file_path}, Last Modified: {last_modified_date}")
                    if last_modified_date > cutoff_date:
                        files_to_check.append(file_path)
                        logger.info(f"Detected file: {file_path}")

    except Exception as e:
        logger.error(f"Error during detection: {e}")

    logger.info(f"Detection completed. Total files checked: {total_files_checked}, Files detected: {len(files_to_check)}")

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

    files_to_convert = []

    try:
        for root, dirs, files in os.walk(directory):
            for file in files:
                total_files_checked += 1
                if any(file.lower().endswith(ft.lower()) for ft in file_types):
                    file_path = os.path.join(root, file)
                    last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                    logger.debug(f"File: {file_path}, Last Modified: {last_modified_date}")
                    if last_modified_date > cutoff_date:
                        files_to_convert.append(file_path)

        for file_path in files_to_convert:
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
class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Legacy file checker and converter")
        self.geometry("920x700")

        self.label_dir = tk.Label(self, text="Target Directory:")
        self.label_dir.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_dir = tk.Entry(self, width=50)
        self.entry_dir.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.button_browse = tk.Button(self, text="Browse", command=self.browse_directory)
        self.button_browse.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.label_days = tk.Label(self, text="Number of Days for Cutoff Date:")
        self.label_days.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_days = tk.Entry(self, width=10)
        self.entry_days.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_delay = tk.Label(self, text="Processing Time per File (seconds):")
        self.label_delay.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_delay = tk.Entry(self, width=10)
        self.entry_delay.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.label_file_types = tk.Label(self, text="File Types:")
        self.label_file_types.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.var_doc = tk.BooleanVar()
        self.check_doc = tk.Checkbutton(self, text="*.doc", variable=self.var_doc)
        self.check_doc.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.var_xls = tk.BooleanVar()
        self.check_xls = tk.Checkbutton(self, text="*.xls", variable=self.var_xls)
        self.check_xls.grid(row=3, column=2, padx=5, pady=5, sticky="w")

        self.label_operation = tk.Label(self, text="Operation:")
        self.label_operation.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.var_operation = tk.StringVar(value="check")
        self.radio_check = tk.Radiobutton(self, text="Check", variable=self.var_operation, value="check")
        self.radio_check.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.radio_convert = tk.Radiobutton(self, text="Convert", variable=self.var_operation, value="convert")
        self.radio_convert.grid(row=4, column=2, padx=5, pady=5, sticky="w")

        self.var_delete_originals = tk.BooleanVar()
        self.check_delete_originals = tk.Checkbutton(self, text="Delete original files after conversion", variable=self.var_delete_originals)
        self.check_delete_originals.grid(row=4, column=3, padx=5, pady=5, sticky="w")

        self.button_run = tk.Button(self, text="Run", command=self.run_operation)
        self.button_run.grid(row=5, column=0, columnspan=4, padx=10, pady=20)

        self.button_help = tk.Button(self, text="Help", command=self.show_help)
        self.button_help.grid(row=5, column=1, columnspan=4, padx=10, pady=20)

        self.text_log = ScrolledText(self, state='disabled', width=110, height=30, font=("Courier", 10))
        self.text_log.grid(row=6, column=0, columnspan=4, padx=10, pady=5)

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_dir.delete(0, tk.END)
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
            detect_files(directory, cutoff_date, file_types)
        elif operation == 'convert':
            convert_files(directory, cutoff_date, delay=delay, file_types=file_types, delete_originals=delete_originals)
        else:
            messagebox.showerror("Error", "Invalid operation. Please select 'check' or 'convert'.")

        messagebox.showinfo("Completed", f"{operation.capitalize()} operation completed successfully.")

    def show_help(self):
        help_window = tk.Toplevel(self)
        help_window.title("Help")
        help_window.geometry("600x300")
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
        tk.Label(help_window, text=help_text, justify="left").pack(padx=10, pady=10)

if __name__ == '__main__':
    app = Application()
    app.mainloop()
