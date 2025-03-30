# -*- coding: utf-8 -*-
"""
Archive Extraction Utility

This script provides a graphical user interface (using Tkinter dialogs)
to select archive files (ZIP, RAR, 7z, TAR, GZ, BZ2) or a folder
containing such files, and extracts each archive into its own
sub-folder within a selected output directory.

It handles potential missing dependencies (patoolib, py7zr) by
attempting to install them using pip. Extraction tasks are run
in parallel using a thread pool for efficiency.

It's recommended to install dependencies beforehand using:
pip install patool py7zr Pillow  # Pillow often needed by tk
Or use the included auto-install feature.
"""

import os
import subprocess
import sys
import zipfile
import logging
import threading
import importlib
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
from typing import List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Configuration ---

# Define supported archive file extensions
SUPPORTED_EXTENSIONS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"}

# List of required packages for core functionality
# Note: 'tk'/'tkinter' is usually built-in. 'zipfile' is built-in.
# 'patool' itself requires other archivers (like rar, 7z) to be installed on the system.
REQUIRED_PACKAGES = {'patoolib': 'patool', 'py7zr': 'py7zr'}

# Maximum number of concurrent extraction threads (adjust as needed)
# Using None lets ThreadPoolExecutor decide based on CPU cores
MAX_WORKERS = os.cpu_count() or 4 # Default to 4 if cpu_count() fails

# --- Logging Setup ---

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(threadName)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout) # Log to console
        # Optionally add FileHandler here:
        # logging.FileHandler("extraction.log")
    ]
)
log = logging.getLogger(__name__)

# --- Dependency Management ---

def check_and_install_packages():
    """Checks if required packages are installed and attempts to install missing ones."""
    packages_installed = True
    for import_name, install_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
            log.debug(f"Package '{import_name}' is already installed.")
        except ImportError:
            log.warning(f"Package '{import_name}' not found. Attempting installation...")
            try:
                # Use check_call to ensure pip command succeeds
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", install_name],
                    stdout=subprocess.DEVNULL, # Hide pip output unless error
                    stderr=subprocess.PIPE
                )
                log.info(f"Successfully installed '{install_name}'.")
                # Verify installation after attempting
                importlib.import_module(import_name)
            except (subprocess.CalledProcessError, ImportError) as e:
                log.error(f"Failed to install package '{install_name}'. Please install it manually.", exc_info=False)
                log.error(f"Error details: {e}")
                messagebox.showerror(
                    "Dependency Error",
                    f"Failed to install required package: '{install_name}'.\n"
                    f"Please install it manually (e.g., 'pip install {install_name}') and restart the script.\n"
                    f"Error: {e}"
                )
                packages_installed = False
            except Exception as e: # Catch other potential errors like permissions
                 log.error(f"An unexpected error occurred during installation of '{install_name}': {e}", exc_info=True)
                 messagebox.showerror(
                    "Installation Error",
                    f"An unexpected error occurred while trying to install '{install_name}'.\n"
                    f"Check permissions and logs.\nError: {e}"
                )
                 packages_installed = False


    # Dynamically import after potential installation
    # This ensures they are available later in the script if installed now
    global patoolib, SevenZipFile
    if packages_installed:
        try:
            import patoolib
            from py7zr import SevenZipFile
        except ImportError:
            # This should ideally not happen if check_and_install worked,
            # but serves as a final safeguard.
            log.critical("Could not import necessary libraries even after installation attempt.")
            messagebox.showerror("Import Error", "Failed to load required libraries. The script cannot continue.")
            packages_installed = False # Ensure we don't proceed

    return packages_installed

# --- Core Extraction Logic ---

def extract_single_archive(archive_path: Path, output_base_dir: Path) -> bool:
    """
    Extracts a single archive file to its own subdirectory within the output base directory.

    Args:
        archive_path: Path to the archive file.
        output_base_dir: The base directory where the extraction subdirectory will be created.

    Returns:
        True if extraction was successful, False otherwise.
    """
    file_ext = archive_path.suffix.lower()
    # Create a dedicated output folder for this archive, named after the archive file (without extension)
    extract_to_dir = output_base_dir / archive_path.stem
    
    try:
        log.info(f"Attempting to extract '{archive_path.name}' to '{extract_to_dir}'...")
        # Ensure the specific output directory exists
        os.makedirs(extract_to_dir, exist_ok=True)

        if file_ext == ".zip":
            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                zip_ref.extractall(extract_to_dir)
        elif file_ext == ".7z":
            # Ensure py7zr is imported (should be handled by check_and_install)
            if 'SevenZipFile' not in globals(): raise ImportError("py7zr library not loaded.")
            with SevenZipFile(archive_path, mode='r') as z:
                z.extractall(path=extract_to_dir)
        elif file_ext in SUPPORTED_EXTENSIONS: # Handles .rar, .tar, .gz, .bz2 via patoolib
             # Ensure patoolib is imported (should be handled by check_and_install)
            if 'patoolib' not in globals(): raise ImportError("patoolib library not loaded.")
            # patoolib requires string paths
            patoolib.extract_archive(str(archive_path), outdir=str(extract_to_dir), verbosity=-1) # -1 for less output
        else:
            # This case should ideally not be reached if filtering is done beforehand
            log.warning(f"Skipping unsupported file format: {archive_path}")
            # Return True because it wasn't an *extraction error*, just unsupported
            return True 

        log.info(f"Successfully extracted: '{archive_path.name}' to '{extract_to_dir}'")
        return True

    except zipfile.BadZipFile:
        log.error(f"Extraction failed: Bad ZIP file format for '{archive_path.name}'.", exc_info=False)
    except ImportError as e:
         log.error(f"Extraction failed for '{archive_path.name}': Missing library dependency. {e}", exc_info=False)
    except Exception as e:
        # Catch patoolib errors (which can be varied) and other generic errors
        log.error(f"Extraction failed for '{archive_path.name}': {e}", exc_info=True) # Log full traceback for debugging
        # Clean up potentially partially created directory if extraction fails midway? Optional.
        # Consider shutil.rmtree(extract_to_dir, ignore_errors=True)

    return False

# --- File Discovery ---

def find_archives_in_folder(input_folder: Path) -> List[Path]:
    """
    Recursively finds all files with supported archive extensions in a folder.

    Args:
        input_folder: The directory to search within.

    Returns:
        A list of Path objects representing the found archive files.
    """
    found_archives = []
    log.info(f"Scanning folder '{input_folder}' for archives...")
    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = Path(root) / file
            if file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                found_archives.append(file_path)
    log.info(f"Found {len(found_archives)} archive(s) in '{input_folder}'.")
    return found_archives

# --- Processing Logic ---

def process_archives(archive_files: List[Path], output_folder: Path) -> Tuple[int, int]:
    """
    Extracts a list of archive files using a thread pool.

    Args:
        archive_files: A list of Path objects for the archives to extract.
        output_folder: The base directory for extractions.

    Returns:
        A tuple containing the count of successful extractions and failed extractions.
    """
    success_count = 0
    failure_count = 0
    total_files = len(archive_files)

    log.info(f"Starting extraction of {total_files} archive(s) using up to {MAX_WORKERS} worker threads...")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS, thread_name_prefix='Extractor') as executor:
        # Submit all extraction tasks
        future_to_path = {executor.submit(extract_single_archive, path, output_folder): path for path in archive_files}

        # Process results as they complete
        for i, future in enumerate(as_completed(future_to_path), 1):
            archive_path = future_to_path[future]
            try:
                success = future.result() # Get result (True/False) from extract_single_archive
                if success:
                    success_count += 1
                    log.info(f"Progress: {i}/{total_files} - Successfully processed '{archive_path.name}'")
                else:
                    failure_count += 1
                    # Error details already logged within extract_single_archive
                    log.warning(f"Progress: {i}/{total_files} - Failed to process '{archive_path.name}'")
            except Exception:
                # Catch unexpected errors from the future itself (should be rare if extract_single_archive handles its exceptions)
                log.exception(f"Progress: {i}/{total_files} - Unexpected error processing future for '{archive_path.name}'")
                failure_count += 1

    log.info(f"Extraction process finished. Success: {success_count}, Failures: {failure_count}")
    return success_count, failure_count

# --- User Interface ---

def run_extraction_ui():
    """Handles user interaction via dialog boxes to select input and output, then starts extraction."""
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window

    # 1. Ask user for input type (Files or Folder)
    choice = messagebox.askyesno(
        title="Select Input Method",
        message="Do you want to select specific archive files?\n\n"
                "• Yes: Choose individual files.\n"
                "• No: Choose a folder containing archives.",
        detail="If you choose 'No', the script will scan the selected folder and its subfolders for archives."
    )

    input_paths: List[Path] = []
    input_description = "" # For the final message

    if choice is None: # User closed the dialog
         log.info("User cancelled input selection. Exiting.")
         messagebox.showinfo("Cancelled", "Operation cancelled by user.")
         return

    if choice:  # User chose "Yes" -> Select Files
        log.info("User chose to select individual files.")
        file_types = [("Archive files", " ".join(f"*{ext}" for ext in SUPPORTED_EXTENSIONS)), ("All files", "*.*")]
        selected_files = filedialog.askopenfilenames(
            title="Select Archive Files to Extract",
            filetypes=file_types
        )
        if not selected_files:
            log.info("No files selected. Exiting.")
            messagebox.showinfo("No Selection", "No files were selected. Exiting.")
            return
        input_paths = [Path(f) for f in selected_files]
        input_description = f"{len(input_paths)} selected file(s)"

    else:  # User chose "No" -> Select Folder
        log.info("User chose to select a folder.")
        input_folder_str = filedialog.askdirectory(title="Select Folder Containing Archives")
        if not input_folder_str:
            log.info("No folder selected. Exiting.")
            messagebox.showinfo("No Selection", "No folder was selected. Exiting.")
            return
        
        input_folder = Path(input_folder_str)
        input_paths = find_archives_in_folder(input_folder)
        input_description = f"folder '{input_folder.name}'"
        if not input_paths:
             log.warning(f"No supported archive files found in '{input_folder}'.")
             messagebox.showwarning("No Archives Found", f"No supported archive files ({', '.join(SUPPORTED_EXTENSIONS)}) were found in the selected folder.")
             return


    # 2. Select Output Folder
    log.info("Prompting for output directory...")
    output_folder_str = filedialog.askdirectory(title="Select Output Folder for Extracted Files")
    if not output_folder_str:
        log.info("No output folder selected. Exiting.")
        messagebox.showinfo("No Selection", "No output folder was selected. Exiting.")
        return
    
    output_folder = Path(output_folder_str)
    log.info(f"Output directory set to: '{output_folder}'")


    # 3. Perform Extraction
    if not input_paths:
        # This case should be handled above, but as a safeguard:
        log.info("No valid input archives to process. Exiting.")
        messagebox.showinfo("Nothing to Do", "No archive files to extract.")
        return

    log.info(f"Starting extraction from {input_description} to '{output_folder}'...")
    success_count, failure_count = process_archives(input_paths, output_folder)

    # 4. Show Summary Report
    summary_title = "Extraction Complete"
    summary_message = f"Extraction process finished for {input_description}.\n\n" \
                      f"Successfully extracted: {success_count} file(s)\n" \
                      f"Failed to extract: {failure_count} file(s)"

    if failure_count > 0:
        summary_title = "Extraction Complete with Errors"
        summary_message += "\n\nPlease check the console or log file for details on failed extractions."
        messagebox.showwarning(summary_title, summary_message)
    else:
        messagebox.showinfo(summary_title, summary_message)

    log.info("Script finished.")


# --- Main Execution ---

if __name__ == "__main__":
    log.info("Script started.")
    # Check dependencies first
    if check_and_install_packages():
        log.info("Dependencies are satisfied.")
        run_extraction_ui()
    else:
        log.error("Dependency check failed. Exiting.")
        # Error message already shown by check_and_install_packages
