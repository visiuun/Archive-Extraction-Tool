# Archive Extraction Tool

This Python script allows you to extract various types of archive files (such as `.zip`, `.rar`, `.7z`, `.tar`, `.gz`, and `.bz2`) to a specified destination folder. You can choose to extract files either from selected individual archives or an entire folder of archives. The tool uses multi-threading to extract archives concurrently for faster processing.

## Features

- **Support for Multiple Archive Formats**:
  - `.zip`
  - `.rar`
  - `.7z`
  - `.tar`
  - `.gz`
  - `.bz2`
- **Multi-threaded Extraction**: Extracts archives concurrently for faster processing.
- **GUI-based File and Folder Selection**: Choose files or folders using a graphical file dialog.
- **Logging and Error Handling**: Logs extraction progress and handles errors with appropriate messages.
- **Package Installation**: The script checks for required packages and installs them if missing.

## Requirements

Ensure the following Python libraries are installed:
- `patool`
- `py7zr`
- `tk`
- `zipfile36`

To install the required packages, the script will automatically attempt to install any missing dependencies when run.

## Installation

1. Ensure you have Python installed on your system.
2. Download the script or clone the repository.
3. Run the script to automatically install any missing dependencies.

## Usage

1. **Run the Script**:
   To run the script, simply execute it with Python:
   ```bash
   python "Bulk zip extractor.py"
   ```

2. **Select Input Type**:
   - The script will ask if you want to select specific archive files or an entire folder. 
   - Choose "Yes" to select individual files or "No" to select a folder containing archive files.

3. **Select Archives**:
   - If you choose files, the file dialog will allow you to select the archive files you want to extract.
   - If you choose a folder, the script will automatically process all the archives in that folder.

4. **Select Output Folder**:
   A dialog will prompt you to choose the destination folder for the extracted files.

5. **Extraction Process**:
   - The script will begin extracting archives.
   - Multiple threads will be created to extract archives concurrently for faster processing.
   - You will receive notifications upon successful completion or errors during extraction.

## Example

1. Run the script:
   ```bash
   python "Bulk zip extractor.py"
   ```

2. Choose to select specific archive files or a folder.
3. Select the desired archives or folder.
4. Choose the output folder where extracted files will be saved.
5. Wait for the extraction to complete.

## Customization

You can modify the script to:
- Support additional archive formats.
- Change the logging behavior to write logs to a file or adjust the logging level.
- Modify how the extraction paths are structured or customize thread behavior.

## Known Limitations

- **File Format**: This script only supports the listed archive formats (`.zip`, `.rar`, `.7z`, `.tar`, `.gz`, `.bz2`). Other formats will trigger a warning.
- **Threading**: The extraction process uses multiple threads which may strain system resources if many large files are processed concurrently.
- **File Overwriting**: If extraction paths contain files with the same name, they will be overwritten without prompt. Customize as needed for specific behavior.

## License

This tool is open-source and free to use. Modify it according to your needs.

---
