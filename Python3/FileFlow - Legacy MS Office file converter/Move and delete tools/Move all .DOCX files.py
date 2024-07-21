import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

def move_docx_files(source_directory, target_directory):
    moved_files = 0
    for root, dirs, files in os.walk(source_directory):
        for file in files:
            if file.lower().endswith('.docx'):
                source_file_path = os.path.join(root, file)
                target_file_path = os.path.join(target_directory, file)
                try:
                    shutil.move(source_file_path, target_file_path)
                    moved_files += 1
                    print(f"Moved: {source_file_path} to {target_file_path}")
                except Exception as e:
                    print(f"Error moving file {source_file_path}: {e}")
    print(f"Total .docx files moved: {moved_files}")

def browse_directory(prompt):
    directory = filedialog.askdirectory(title=prompt)
    if directory:
        return directory
    else:
        messagebox.showerror("Error", f"No directory selected for {prompt.lower()}.")
        return None

def run_move_operation():
    source_directory = browse_directory("Select Source Directory")
    if not source_directory:
        return
    target_directory = browse_directory("Select Target Directory")
    if not target_directory:
        return
    
    if source_directory == target_directory:
        messagebox.showerror("Error", "Source and target directories cannot be the same.")
        return

    move_docx_files(source_directory, target_directory)
    messagebox.showinfo("Completed", "Move operation completed successfully.")

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    run_move_operation()
