import os
import tkinter as tk
from tkinter import filedialog, messagebox

def delete_docx_files(directory):
    deleted_files = 0
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.docx'):
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                    deleted_files += 1
                    print(f"Deleted: {file_path}")
                except Exception as e:
                    print(f"Error deleting file {file_path}: {e}")
    print(f"Total .docx files deleted: {deleted_files}")

def browse_directory():
    directory = filedialog.askdirectory()
    if directory:
        return directory
    else:
        messagebox.showerror("Error", "No directory selected.")
        return None

def run_deletion():
    directory = browse_directory()
    if directory:
        delete_docx_files(directory)
        messagebox.showinfo("Completed", "Deletion operation completed successfully.")

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    run_deletion()
