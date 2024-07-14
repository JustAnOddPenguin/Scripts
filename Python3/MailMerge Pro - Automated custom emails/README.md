# MailMerge Pro

## Description

MailMerge Pro is a Python-based application designed to automate the process of sending emails with PDF attachments to multiple recipients. The program reads recipient information from a CSV file and attaches relevant PDF files based on matching IDs.

## Features

- User-friendly GUI for selecting CSV and PDF files
- Logs all activities to a file for easy debugging and record-keeping
- Handles various email sending errors gracefully
- Clears input fields after successful email dispatch

## Prerequisites

Before running MailMerge Pro, ensure you have the following:

- Python 3.x installed on your machine
- Required Python libraries: `tkinter`, `logging`, `csv`, `smtplib`, `ssl`, and `email`
- A valid SMTP server to send emails

## Installation

1. Clone the repository to your local machine:

    ```bash
    git clone https://github.com/JustAnOddPenguin/Scripts.git
    cd mailmerge-pro
    ```
    
## Usage

1. **Configure SMTP Settings:**

    Open the `MailMerge Pro.py` file and update the following lines with your SMTP server details:

    ```python
    smtp_server = "your_smtp_server"
    smtp_port = "your_smtp_port"
    ```

2. **Run the Application:**

    Execute the following command to start the GUI:

    ```bash
    python MailMerge Pro.py
    ```

3. **Using the GUI:**

    - **Sender Email:** Enter the email address from which you want to send emails.
    - **Email Subject:** Enter the subject of the email.
    - **Email Message:** Enter the body of the email.
    - **Browse PDF Folder:** Select the folder containing the PDF files.
    - **Browse CSV:** Select the CSV file with recipient information.
    - **Send:** Click the "Send" button to start sending emails.

4. **CSV File Format:**

    Ensure your CSV file follows this format:

    ```csv
    invoice_id,email1,email2,...
    12345,example1@example.com,example2@example.com
    67890,example3@example.com
    ```

## Logging

Logs are saved in a `log` folder in the same directory where the script is run. The log file is named `Csv_Email_Log.txt`. 

