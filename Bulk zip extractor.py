import os
import subprocess
import sys
import zipfile
import patoolib
from py7zr import SevenZipFile
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
import logging
import threading

# List of required packages
required_packages = ['patool', 'py7zr', 'tk', 'zipfile36']

def install_required_packages():
    # Ensure that each required package is installed
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Package {package} not found. Installing...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Install required packages before running the rest of the script
install_required_packages()

# Setup logging for better debugging and reporting
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_archive(file_path, output_folder):
    file_ext = file_path.suffix.lower()
    output_folder = Path(output_folder)  # Ensure output_folder is a Path object
    extract_to_path = output_folder / file_path.stem  # Combine paths correctly
    
    # Ensure the output folder for this archive exists
    os.makedirs(extract_to_path, exist_ok=True)

    try:
        if file_ext == ".zip":
            # Extract ZIP files
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_to_path)
                logging.info(f"Extracted: {file_path} to {extract_to_path}")
        elif file_ext in [".rar", ".tar", ".gz", ".bz2"]:
            # Extract other archive formats using patool
            patoolib.extract_archive(str(file_path), outdir=str(extract_to_path))
            logging.info(f"Extracted: {file_path} to {extract_to_path}")
        elif file_ext == ".7z":
            # Extract 7z files using py7zr
            with SevenZipFile(file_path, mode='r') as z:
                z.extractall(path=extract_to_path)
                logging.info(f"Extracted: {file_path} to {extract_to_path}")
        else:
            logging.warning(f"Unsupported file format: {file_path}")
            messagebox.showwarning("Unsupported Format", f"Unsupported file format: {file_path}")
    except zipfile.BadZipFile:
        logging.error(f"Bad ZIP file: {file_path}")
        messagebox.showerror("Extraction Error", f"Failed to extract: {file_path}. Bad ZIP file.")
    except patoolib.util.PatoolError:
        logging.error(f"Failed to extract with Patool: {file_path}")
        messagebox.showerror("Extraction Error", f"Failed to extract: {file_path} with Patool.")
    except Exception as e:
        logging.error(f"Failed to extract {file_path}: {e}")
        messagebox.showerror("Extraction Error", f"Failed to extract: {file_path}\nError: {str(e)}")

def extract_archives_from_folder(input_folder, output_folder):
    # Traverse the input folder recursively and create a thread for each file
    threads = []
    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = Path(root) / file
            if file_path.suffix.lower() in [".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"]:
                # Create a thread for each extraction task
                thread = threading.Thread(target=extract_archive, args=(file_path, output_folder))
                threads.append(thread)
                thread.start()

    # Wait for all threads to finish
    for thread in threads:
        thread.join()

def extract_selected_files(files, output_folder):
    threads = []
    for file in files:
        file_path = Path(file)
        if file_path.suffix.lower() in [".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"]:
            # Create a thread for each extraction task
            thread = threading.Thread(target=extract_archive, args=(file_path, output_folder))
            threads.append(thread)
            thread.start()

    # Wait for all threads to finish
    for thread in threads:
        thread.join()

def select_input_and_extract():
    root = Tk()
    root.withdraw()  # Hide the root window
    
    # Ask the user to choose files or a folder
    choice = messagebox.askyesno(
        title="Select Input Type",
        message="Would you like to select specific archive files instead of an entire folder?\n\nYes: Select Files\nNo: Select Folder"
    )
    
    if choice:  # User chooses files
        input_files = filedialog.askopenfilenames(
            title="Select archive files",
            filetypes=[("Archive files", "*.zip *.rar *.7z *.tar *.gz *.bz2"), ("All files", "*.*")]
        )
        if not input_files:
            logging.info("No files selected. Exiting...")
            messagebox.showinfo("No Files", "No files selected. Exiting...")
            return
    else:  # User chooses folder
        input_folder = filedialog.askdirectory(title="Select the input folder containing archive files")
        if not input_folder:
            logging.info("No folder selected. Exiting...")
            messagebox.showinfo("No Folder", "No folder selected. Exiting...")
            return
        input_files = None

    # Prompt user to select the output folder
    output_folder = filedialog.askdirectory(title="Select the output folder to extract archive files to")
    if not output_folder:
        logging.info("No output folder selected. Exiting...")
        messagebox.showinfo("No Output Folder", "No output folder selected. Exiting...")
        return

    # Extract based on the user's choice
    if input_files:
        extract_selected_files(input_files, output_folder)
    else:
        extract_archives_from_folder(input_folder, output_folder)

    messagebox.showinfo("Extraction Completed", "Extraction completed successfully!")

# Run the input selection and archive extraction process
if __name__ == "__main__":
    select_input_and_extract()