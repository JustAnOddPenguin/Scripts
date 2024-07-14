import os
import logging
import colorlog
from datetime import datetime, timedelta

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
def configure_logging(directory, suffix=' - Check'):
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

def detect_files(directory, cutoff_date, file_types=['.xls', '.doc']):
    logger = logging.getLogger()
    logger.info(f"Starting detection in directory: {directory}")
    logger.info(f"Cutoff date for file modification: {cutoff_date.strftime('%Y-%m-%d')}")

    try:
        # Collect all specified file types in the directory
        files_to_check = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if any(file.endswith(ft) for ft in file_types):
                    file_path = os.path.join(root, file)
                    # Check if the file was modified after the cutoff date
                    last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if last_modified_date > cutoff_date:
                        files_to_check.append(file_path)
                        logger.info(f"Detected file: {file_path}")

    except Exception as e:
        logger.error(f"Error during detection: {e}")

    logger.info("Detection completed.")

# Example usage
directory = r'C:\Temp\Target\Attempt4'  # Replace with your directory
file_types = ['.xls', '.doc']  # Specify the file types to detect
cutoff_date = datetime.now() - timedelta(days=60)  # Set the cutoff date for file modification

configure_logging(directory, ' - Check')
detect_files(directory, cutoff_date, file_types)
