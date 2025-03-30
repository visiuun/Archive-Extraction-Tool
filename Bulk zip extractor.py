# -*- coding: utf-8 -*-
"""
Archive Extraction Utility

This script provides a graphical user interface (using Tkinter dialogs)
to select archive files (ZIP, RAR, 7z, TAR, GZ, BZ2) or a folder
containing such files, and extracts each archive into its own
sub-folder within a selected output directory.

Includes an option to recursively extract archives found within
extracted content.

Handles potential missing dependencies (patoolib, py7zr) by
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
# Removed threading import as we use concurrent.futures
import importlib
import shutil  # Added for potential cleanup later, though not used yet
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
from typing import List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Configuration ---

# Define supported archive file extensions
SUPPORTED_EXTENSIONS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"}

# List of required packages for core functionality
REQUIRED_PACKAGES = {'patoolib': 'patool', 'py7zr': 'py7zr'}

# Maximum number of concurrent extraction threads
MAX_WORKERS = os.cpu_count() or 4

# --- Logging Setup ---

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(threadName)s] - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
        # logging.FileHandler("extraction.log")
    ]
)
log = logging.getLogger(__name__)

# --- Dependency Management ---

# [check_and_install_packages function remains the same - omitted for brevity]
def check_and_install_packages():
    """Checks if required packages are installed and attempts to install missing ones."""
    packages_installed = True
    global patoolib, SevenZipFile # Ensure these are global
    
    # Check Tkinter availability early
    try:
        import tkinter
    except ImportError:
        log.critical("Tkinter library not found. This script requires Tkinter for its UI.")
        # No messagebox here as Tkinter itself is missing
        print("ERROR: Tkinter library not found. Cannot run the graphical interface.", file=sys.stderr)
        return False
        
    for import_name, install_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
            log.debug(f"Package '{import_name}' is already installed.")
        except ImportError:
            log.warning(f"Package '{import_name}' not found. Attempting installation via pip...")
            try:
                # Use check_call to ensure pip command succeeds; hide output unless error
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", install_name],
                    stdout=subprocess.DEVNULL, stderr=subprocess.PIPE
                )
                log.info(f"Successfully installed '{install_name}'. Verifying import...")
                # Verify installation after attempting
                importlib.import_module(import_name)
                log.info(f"Successfully imported '{import_name}' after installation.")
            except (subprocess.CalledProcessError) as e:
                pip_error = e.stderr.decode() if e.stderr else str(e)
                log.error(f"Failed to install package '{install_name}' using pip.", exc_info=False)
                log.error(f"Pip error details: {pip_error}")
                messagebox.showerror(
                    "Dependency Installation Error",
                    f"Failed to install required package: '{install_name}'.\n"
                    f"Please install it manually (e.g., 'pip install {install_name}') and restart the script.\n"
                    f"Error: {pip_error}"
                )
                packages_installed = False
            except ImportError as e:
                 log.error(f"Failed to import package '{install_name}' even after attempting installation.", exc_info=False)
                 messagebox.showerror(
                    "Dependency Import Error",
                    f"Failed to import required package: '{install_name}' after installation.\n"
                    f"There might be an issue with the installation or Python environment.\nError: {e}"
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

    # Attempt to dynamically import after potential installation
    if packages_installed:
        try:
            # These imports might fail if installation didn't actually succeed
            # despite check_call returning 0 (rare, but possible)
            import patoolib
            from py7zr import SevenZipFile
            # Assign to globals explicitly if needed elsewhere, but check_and_install should be called first
            globals()['patoolib'] = patoolib
            globals()['SevenZipFile'] = SevenZipFile
            log.debug("Successfully imported core extraction libraries.")
        except ImportError as e:
            log.critical(f"Could not import necessary libraries even after installation check: {e}")
            messagebox.showerror("Import Error", f"Failed to load required libraries ({e}). The script cannot continue.")
            packages_installed = False # Ensure we don't proceed

    return packages_installed


# --- Core Extraction Logic ---

def extract_single_archive(
    archive_path: Path,
    output_base_dir: Path,
    extract_recursively: bool = False # New parameter
) -> bool:
    """
    Extracts a single archive file to its own subdirectory and optionally
    extracts nested archives found within.

    Args:
        archive_path: Path to the archive file.
        output_base_dir: The base directory where the extraction subdirectory
                         for *this* archive will be created.
        extract_recursively: If True, scan extracted contents for more archives
                             and extract them as well.

    Returns:
        True if the primary extraction of `archive_path` was successful,
        False otherwise. Nested extraction failures are logged but do not
        change the return value for the parent.
    """
    if not archive_path.is_file():
        log.error(f"Input path is not a file: {archive_path}")
        return False
        
    file_ext = archive_path.suffix.lower()
    if file_ext not in SUPPORTED_EXTENSIONS:
        log.warning(f"Skipping unsupported file format: {archive_path}")
        return True # Not an error, just skipped

    # Create a dedicated output folder for this archive
    # e.g., archive.zip -> output_base_dir/archive/
    extract_to_dir = output_base_dir / archive_path.stem
    
    primary_extraction_succeeded = False
    try:
        log.info(f"Attempting to extract '{archive_path.name}' -> '{extract_to_dir}'")
        os.makedirs(extract_to_dir, exist_ok=True)

        if file_ext == ".zip":
            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                zip_ref.extractall(extract_to_dir)
        elif file_ext == ".7z":
            if 'SevenZipFile' not in globals(): raise ImportError("py7zr library not loaded.")
            with SevenZipFile(archive_path, mode='r') as z:
                z.extractall(path=extract_to_dir)
        elif file_ext in SUPPORTED_EXTENSIONS: # Handles .rar, .tar, .gz, .bz2
            if 'patoolib' not in globals(): raise ImportError("patoolib library not loaded.")
            patoolib.extract_archive(str(archive_path), outdir=str(extract_to_dir), verbosity=-1)
        # No 'else' needed due to check at the start

        log.info(f"Successfully extracted primary archive: '{archive_path.name}'")
        primary_extraction_succeeded = True

    except zipfile.BadZipFile:
        log.error(f"Extraction failed: Bad ZIP file format for '{archive_path.name}'.")
    except ImportError as e:
         log.error(f"Extraction failed for '{archive_path.name}': Missing library dependency. {e}")
    except FileNotFoundError:
         log.error(f"Extraction failed: Archive file not found at '{archive_path}'.")
    except PermissionError:
         log.error(f"Extraction failed: Permission denied for '{archive_path}' or '{extract_to_dir}'.")
    except Exception as e:
        # Catch patoolib errors and other generic exceptions
        # PatoolError might need specific catching if available: from patoolib.util import PatoolError
        log.error(f"Extraction failed for '{archive_path.name}': {e}", exc_info=True)
        # Consider cleaning up partially created directory:
        # if extract_to_dir.exists(): shutil.rmtree(extract_to_dir, ignore_errors=True)

    # --- Recursive Extraction Logic ---
    if primary_extraction_succeeded and extract_recursively:
        log.info(f"Recursively scanning extracted folder: '{extract_to_dir}'")
        nested_archives_found = []
        try:
            for item in extract_to_dir.rglob('*'): # rglob searches recursively
                if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS:
                    nested_archives_found.append(item)
        except Exception as e:
             log.error(f"Error scanning for nested archives in '{extract_to_dir}': {e}", exc_info=True)

        if nested_archives_found:
            log.info(f"Found {len(nested_archives_found)} nested archive(s) in '{extract_to_dir}'. Processing...")
            for nested_archive_path in nested_archives_found:
                # Extract nested archives *into the same directory* where they were found
                # Pass True to continue recursion if needed.
                log.info(f"--- Starting nested extraction for: '{nested_archive_path.relative_to(output_base_dir)}'")
                nested_success = extract_single_archive(
                    nested_archive_path,
                    nested_archive_path.parent, # Output base is the parent dir of the nested archive
                    extract_recursively=True
                )
                if nested_success:
                    log.info(f"--- Finished nested extraction for: '{nested_archive_path.name}' (Success)")
                    # Optional: Delete the nested archive file after successful extraction?
                    # try:
                    #     nested_archive_path.unlink()
                    #     log.info(f"--- Deleted nested archive file: '{nested_archive_path.name}'")
                    # except OSError as e:
                    #     log.warning(f"--- Could not delete nested archive '{nested_archive_path.name}': {e}")
                else:
                    log.warning(f"--- Finished nested extraction for: '{nested_archive_path.name}' (FAILED)")
        else:
            log.info(f"No nested archives found in '{extract_to_dir}'.")

    return primary_extraction_succeeded

# --- File Discovery ---

# [find_archives_in_folder function remains the same - omitted for brevity]
def find_archives_in_folder(input_folder: Path) -> List[Path]:
    """
    Recursively finds all files with supported archive extensions in a folder.

    Args:
        input_folder: The directory to search within.

    Returns:
        A list of Path objects representing the found archive files.
    """
    found_archives = []
    log.info(f"Scanning folder '{input_folder}' for *top-level* archives...")
    # Note: This only finds archives in the initial scan path.
    # Recursive extraction handles archives *inside* other archives later.
    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = Path(root) / file
            if file_path.is_file() and file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                found_archives.append(file_path)
    log.info(f"Found {len(found_archives)} top-level archive(s) in '{input_folder}'.")
    return found_archives


# --- Processing Logic ---

def process_archives(
    archive_files: List[Path],
    output_folder: Path,
    extract_recursively: bool # New parameter
) -> Tuple[int, int]:
    """
    Extracts a list of top-level archive files using a thread pool.
    Optionally triggers recursive extraction within each thread.

    Args:
        archive_files: A list of Path objects for the top-level archives.
        output_folder: The base directory for extractions.
        extract_recursively: Passed to each extraction task.

    Returns:
        A tuple containing the count of successful and failed *top-level* extractions.
    """
    success_count = 0
    failure_count = 0
    total_files = len(archive_files)

    log.info(f"Starting extraction of {total_files} top-level archive(s) using up to {MAX_WORKERS} workers.")
    if extract_recursively:
        log.info("Recursive extraction is ENABLED.")
    else:
        log.info("Recursive extraction is DISABLED.")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS, thread_name_prefix='Extractor') as executor:
        # Submit all top-level extraction tasks, passing the recursive flag
        future_to_path = {
            executor.submit(extract_single_archive, path, output_folder, extract_recursively): path
            for path in archive_files
        }

        # Process results as they complete
        for i, future in enumerate(as_completed(future_to_path), 1):
            archive_path = future_to_path[future]
            try:
                # Result indicates success/failure of the *primary* extraction only
                success = future.result()
                if success:
                    success_count += 1
                    log.info(f"Progress: {i}/{total_files} - Successfully processed top-level '{archive_path.name}'")
                else:
                    failure_count += 1
                    log.warning(f"Progress: {i}/{total_files} - Failed to process top-level '{archive_path.name}'")
            except Exception as e:
                # Catch unexpected errors from the future itself
                log.exception(f"Progress: {i}/{total_files} - Unexpected error processing future for '{archive_path.name}': {e}")
                failure_count += 1

    log.info(f"Top-level extraction process finished. Success: {success_count}, Failures: {failure_count}")
    return success_count, failure_count

# --- User Interface ---

def run_extraction_ui():
    """Handles user interaction via dialog boxes to select input, output, options, then starts extraction."""
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window

    # 1. Ask user for input type (Files or Folder)
    # [Input selection logic remains the same - omitted for brevity]
    choice = messagebox.askyesno(
        title="Select Input Method",
        message="Do you want to select specific archive files?\n\n"
                "• Yes: Choose individual files.\n"
                "• No: Choose a folder containing archives.",
        detail="If 'No', the script scans the selected folder and subfolders for archives."
    )

    input_paths: List[Path] = []
    input_description = "" 

    if choice is None: 
         log.info("User cancelled input selection. Exiting.")
         messagebox.showinfo("Cancelled", "Operation cancelled by user.")
         return

    if choice:  # Select Files
        log.info("User chose to select individual files.")
        file_types = [("Archive files", " ".join(f"*{ext}" for ext in SUPPORTED_EXTENSIONS)), ("All files", "*.*")]
        selected_files = filedialog.askopenfilenames(
            title="Select Archive Files to Extract", filetypes=file_types
        )
        if not selected_files:
            log.info("No files selected. Exiting.")
            messagebox.showinfo("No Selection", "No files were selected. Exiting.")
            return
        input_paths = [Path(f).resolve() for f in selected_files] # Use resolved paths
        input_description = f"{len(input_paths)} selected file(s)"
    else:  # Select Folder
        log.info("User chose to select a folder.")
        input_folder_str = filedialog.askdirectory(title="Select Folder Containing Archives")
        if not input_folder_str:
            log.info("No folder selected. Exiting.")
            messagebox.showinfo("No Selection", "No folder was selected. Exiting.")
            return
        
        input_folder = Path(input_folder_str).resolve() # Use resolved path
        input_paths = find_archives_in_folder(input_folder)
        input_description = f"folder '{input_folder.name}'"
        if not input_paths:
             log.warning(f"No supported top-level archive files found in '{input_folder}'.")
             messagebox.showwarning("No Archives Found", f"No supported top-level archive files ({', '.join(SUPPORTED_EXTENSIONS)}) were found in the selected folder.")
             # Allow proceeding even if no top-level found, in case user selected files earlier somehow? No, exit here.
             return

    # 2. Select Output Folder
    # [Output selection logic remains the same - omitted for brevity]
    log.info("Prompting for output directory...")
    output_folder_str = filedialog.askdirectory(title="Select Output Folder for Extracted Files")
    if not output_folder_str:
        log.info("No output folder selected. Exiting.")
        messagebox.showinfo("No Selection", "No output folder was selected. Exiting.")
        return
    output_folder = Path(output_folder_str).resolve() # Use resolved path
    log.info(f"Output directory set to: '{output_folder}'")
    
    # Prevent extracting into the source folder if a folder was selected
    if not choice and input_folder == output_folder:
         log.error("Output folder cannot be the same as the input folder.")
         messagebox.showerror("Invalid Selection", "The output folder cannot be the same as the input folder when processing a whole folder.")
         return
    # Could add a check for output being *inside* input too

    # 3. Ask about Recursive Extraction (NEW)
    log.info("Asking user about recursive extraction.")
    perform_recursive_extraction = messagebox.askyesno(
        title="Recursive Extraction",
        message="Extract archives found inside other archives?\n\n"
                "• Yes: Scan extracted folders and extract any nested archives.\n"
                "• No: Only extract the top-level selected archives.",
        detail="Example: If archive.zip contains nested.rar, choosing 'Yes' will extract both."
    )
    if perform_recursive_extraction is None: # User closed dialog
        log.info("Recursive option selection cancelled. Exiting.")
        messagebox.showinfo("Cancelled", "Operation cancelled by user.")
        return

    # 4. Perform Extraction
    if not input_paths:
        log.info("No valid input archives to process. Exiting.")
        messagebox.showinfo("Nothing to Do", "No archive files to extract.")
        return

    log.info(f"Starting extraction from {input_description} to '{output_folder}'...")
    success_count, failure_count = process_archives(
        input_paths, output_folder, perform_recursive_extraction # Pass the flag
    )

    # 5. Show Summary Report
    summary_title = "Extraction Complete"
    recursive_status = "enabled" if perform_recursive_extraction else "disabled"
    summary_message = f"Extraction process finished for {input_description}.\n" \
                      f"(Recursive extraction was {recursive_status})\n\n" \
                      f"Successfully extracted (top-level): {success_count} file(s)\n" \
                      f"Failed to extract (top-level): {failure_count} file(s)"

    if failure_count > 0:
        summary_title = "Extraction Complete with Errors"
        summary_message += "\n\nPlease check the console or log file for details on failed extractions (including nested ones if enabled)."
        messagebox.showwarning(summary_title, summary_message)
    else:
        messagebox.showinfo(summary_title, summary_message)

    log.info("Script finished.")


# --- Main Execution ---

if __name__ == "__main__":
    log.info("Archive Extraction Script started.")
    # Ensure Tkinter root is created early for messageboxes during init
    try:
        root = Tk()
        root.withdraw()
    except Exception as e:
        log.critical(f"Failed to initialize Tkinter: {e}", exc_info=True)
        print(f"ERROR: Failed to initialize GUI environment: {e}", file=sys.stderr)
        sys.exit(1)

    # Check dependencies first
    if check_and_install_packages():
        log.info("Dependencies seem satisfied.")
        try:
             run_extraction_ui()
        except Exception as e:
             log.critical("An unexpected error occurred during the extraction process.", exc_info=True)
             messagebox.showerror("Runtime Error", f"An critical error occurred:\n{e}\n\nPlease check the logs.")
    else:
        log.error("Dependency check failed. Exiting.")
        # Error message already shown by check_and_install_packages

    # Clean up Tkinter root if it exists
    try:
        root.destroy()
    except: # nosec
        pass 
    log.info("Script exit.")
