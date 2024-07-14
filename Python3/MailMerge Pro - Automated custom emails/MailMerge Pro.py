import logging
import os
import sys
import csv
import pathlib
import smtplib, ssl
import tkinter as tk
import tkinter.messagebox as messagebox
import json

from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Create a logger object
logger = logging.getLogger(__name__)
# Set the logging level
logger.setLevel(logging.DEBUG)

# Determine the directory where the script is being run
current_dir = os.path.dirname(os.path.abspath(__file__))

# Create a log folder if it doesn't exist

log_dir = os.path.join(current_dir, 'log')
os.makedirs(log_dir, exist_ok=True)

# Create a file handler
log_file_path = os.path.join(log_dir, 'Csv_Email_Log.txt')
handler = logging.FileHandler(log_file_path)

# Create a log formatter
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

# Add the formatter to the file handler
handler.setFormatter(formatter)

# Add the file handler to the logger
logger.addHandler(handler)

class EmailGUI:
    def __init__(self, master):
        self.master = master
        self.csv_file = None
        self.folder = None
        self.sender_email = None
        
        # Create a label and textbox for the sender email
        self.sender_email_label = tk.Label(master, text="Sender email:")
        self.sender_email_label.pack()
        self.sender_email_text = tk.Text(master, height=1, width=40)
        self.sender_email_text.pack()

        # Create a label and textbox for the email subject
        self.subject_label = tk.Label(master, text="Email subject:")
        self.subject_label.pack()
        self.subject_text = tk.Text(master, height=1, width=40)
        self.subject_text.pack()
        
        # Create a label and textbox for the email description
        self.description_label = tk.Label(master, text="Email message:")
        self.description_label.pack()
        self.description_text = tk.Text(master, height=15, width=40)
        self.description_text.pack()

        # Add a button for browsing for a folder
        self.browse_button = tk.Button(master, text="Browse PDF folder", command=self.browse_folder)
        self.browse_button.pack()
        
        # Add a label to display the selected folder
        self.folder_label = tk.Label(master, text="No folder selected")
        self.folder_label.pack()

        # Add a button for browsing for a CSV file
        self.browse_csv_button = tk.Button(master, text="Browse CSV", command=self.browse_csv)
        self.browse_csv_button.pack()

        # Add a label to display the selected folder
        self.csv_label = tk.Label(master, text="No csv selected")
        self.csv_label.pack()

        # Add a button for sending
        self.send_button = tk.Button(master, text="Send", command=self.send_email)
        self.send_button.pack()

    def browse_csv(self):
        # Open a file selection dialog and select a CSV file
        self.csv_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.csv_label.config(text="csv selected")
       
    def browse_folder(self):
        # Open a folder selection dialog and update the label with the selected folder
        folder = filedialog.askdirectory()
        self.folder_label.config(text=folder)
        self.folder = folder

    def pdf_files(self):
        # Get the list of PDF files in the selected folder
        return [file for file in os.listdir(self.folder) if file.endswith(".pdf")]

    def send_email(self):
        # Get sender, subject and description
        self.sender_email = self.sender_email_text.get("1.0", "end-1c")
        # self.sender_email = sender_email
        subject = self.subject_text.get("1.0", "end-1c")
        
        if not self.sender_email:
            messagebox.showerror("Error", "Please enter a sender email")
            return
        if not subject:
            messagebox.showerror("Error", "Please enter an email subject")
            return

        # SMTP SERVER and PORT
        smtp_server = "SMTP-Sever"
        smtp_port = 25

        # Server connection
        try:
            # Create SMTP server
            server = smtplib.SMTP(smtp_server, smtp_port)
            # Error handling
        except smtplib.SMTPConnectError as e:
            logger.exception("An error occurred while connecting to the server")
            messagebox.showerror("Error", "An error occurred while connecting to the server")
        except smtplib.SMTPAuthenticationError as e:
            logger.exception("Authentication error")
            messagebox.showerror("Error", "Authentication error")
        except smtplib.SMTPHeloError as e:
            logger.exception("Error sending HELO message")
            messagebox.showerror("Error", "Error sending HELO message")
        except smtplib.SMTPRecipientsRefused as e:
            logger.exception("At least one recipient was refused")
            messagebox.showerror("Error", "At least one recipient was refused")
        except smtplib.SMTPSenderRefused as e:
            logger.exception("Sender address was refused")
            messagebox.showerror("Error", "Sender address was refused")
        except smtplib.SMTPDataError as e:
            logger.exception("Error sending data")
            messagebox.showerror("Error", "Error sending data")
        except Exception as e:
            logger.exception("An error occurred while creating the SSL context")
            messagebox.showerror("Error", "An error occurred while creating the SSL context")
  
        # Open CSV file # Find matching record for this file's ID
        with open(self.csv_file, "r") as csv_file:
            reader = csv.reader(csv_file,delimiter=",")
            next(reader) # Skips the header row
            for row in reader:
                # Get Invoice name excluding the .pdf suffix
                for pdf in self.pdf_files():   
                    invoice_ID = pdf[0:-4]
                    if row[0] != invoice_ID:
                        continue
                    csv_name, csv_email = row[0], row[1:]
                    file_path = os.path.join(self.folder, f"{invoice_ID}.pdf")
                    if csv_name == invoice_ID:   
                        #Found the record, begin to send email
                        for recipient_email in csv_email:
                            # Construct the email
                            message = MIMEMultipart()
                            message["From"] = self.sender_email
                            message["To"] = recipient_email
                            message["Subject"] = self.subject_text.get("1.0", "end-1c")
                            message.attach(MIMEText(self.description_text.get("1.0", "end-1c"), "plain"))
                            
                            # Check if the PDF file exists
                            if os.path.exists(file_path):
                                # Create the attachment
                                attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="pdf")
                                attachment.add_header("Content-Disposition", "attachment", filename=f"{invoice_ID}.pdf")
                                # Add the attachment to the email
                                message.attach(attachment)
                            else:
                                # Log an error if the PDF file doesn't exist
                                logger.error(f"PDF file not found for invoice ID {invoice_ID}")
                           
                            # Send the email
                            try:
                                print(f"Sending {invoice_ID} to email {recipient_email}")
                                logger.info("Email sending to {}".format(recipient_email))
                                server.sendmail(self.sender_email, recipient_email, message.as_string())
                                print(f"Email {subject} sent {invoice_ID} document to {recipient_email}")
                                logger.info(f"Email {subject} sent {invoice_ID} document to {recipient_email}")
                        
                            except smtplib.SMTPRecipientsRefused as e:
                                print("Failed to send email: invalid recipient address")
                            except smtplib.SMTPDataError as e:
                                print("Failed to send email: data error")
                            except smtplib.SMTPHeloError as e:
                                print("Failed to send email: invalid HELO/EHLO message")
                            except smtplib.SMTPSenderRefused as e:
                                print("Failed to send email: invalid sender address")
                            except smtplib.SMTPAuthenticationError as e:
                                print("Failed to send email: authentication error")
                            except smtplib.SMTPException as e:
                                print("Failed to send email: general error")      
        # Close the connection to the server
        server.quit()
        messagebox.showinfo("Success", "Emails sent successfully")
        logger.info("Success, Emails sent successfully")
        # Clear the text boxes
        self.sender_email_text.delete("1.0", "end")
        self.subject_text.delete("1.0", "end")
        self.description_text.delete("1.0", "end")
        self.folder_label.config(text="No folder selected")
        self.csv_label.config(text="No csv selected")

def main():
    # Set up the GUI
    root = tk.Tk()
    root.title('.CSV emailer')
    app = EmailGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
