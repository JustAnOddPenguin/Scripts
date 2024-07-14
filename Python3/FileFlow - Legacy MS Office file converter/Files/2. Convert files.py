import os
import time
import xlwings as xw
import logging
import colorlog
from tqdm import tqdm
from datetime import datetime, timedelta
from docx import Document
import win32com.client as win32

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
def configure_logging(directory, suffix=' - Convert'):
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

    # Console handler with color
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)

    # File handler
    file_handler = logging.FileHandler(log_file_name, mode='w')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

def convert_files(directory, cutoff_date, delay=2, file_types=['.xls', '.doc']):
    logger = logging.getLogger()
    logger.info(f"Starting conversion in directory: {directory}")
    logger.info(f"Cutoff date for file modification: {cutoff_date.strftime('%Y-%m-%d')}")

    # Create an instance of the Excel application
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    word = win32.Dispatch("Word.Application")
    word.Visible = False

    try:
        # Collect all specified file types in the directory
        files_to_convert = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if any(file.endswith(ft) for ft in file_types):
                    file_path = os.path.join(root, file)
                    # Check if the file was modified within the last 2 months
                    last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if last_modified_date > cutoff_date:
                        files_to_convert.append(file_path)
        
        # Create a progress bar
        for file_path in tqdm(files_to_convert, desc="Processing files", unit="file"):
            logger.info(f"Processing file: {file_path}")

            try:
                if file_path.endswith('.xls'):
                    # Open the .xls file with xlwings
                    wb = app.books.open(file_path)
                    logger.info(f"Opened file: {file_path}")

                    # Create the new file path with .xlsx extension
                    new_file_path = file_path.rsplit('.', 1)[0] + '.xlsx'
                    logger.info(f"Converting to: {new_file_path}")

                    # Save the workbook with the .xlsx extension
                    wb.save(new_file_path)
                    logger.info(f"Saved file as: {new_file_path}")

                    # Close the workbook
                    wb.close()
                    logger.info(f"Closed file: {file_path}")

                elif file_path.endswith('.doc'):
                    # Open the .doc file with pywin32
                    doc = word.Documents.Open(file_path)
                    logger.info(f"Opened file: {file_path}")

                    # Create the new file path with .docx extension
                    new_file_path = file_path.rsplit('.', 1)[0] + '.docx'
                    logger.info(f"Converting to: {new_file_path}")

                    # Save the document with the .docx extension
                    doc.SaveAs(new_file_path, FileFormat=16)  # 16 corresponds to the wdFormatXMLDocument file format
                    doc.Close()
                    logger.info(f"Saved file as: {new_file_path}")

                # Optionally, you can delete the old .xls or .doc file
                os.remove(file_path)
                logger.info(f"Deleted original file: {file_path}")

            except Exception as e:
                logger.error(f"Error processing file {file_path}: {e}")

            # Wait for the specified delay before processing the next file
            time.sleep(delay)

    finally:
        app.quit()
        word.Quit()

    logger.info("Conversion completed.")

# Example usage
directory = r'C:\Temp\Target\Attempt4'  # Replace with your directory
file_types = ['.xls', '.doc']  # Specify the file types to convert
cutoff_date = datetime.now() - timedelta(days=60)  # Set the cutoff date for file modification
delay_seconds = 0.5  # Delay in seconds between processing files

configure_logging(directory, ' - Convert')
convert_files(directory, cutoff_date, delay=delay_seconds, file_types=file_types)
