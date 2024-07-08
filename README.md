# Network Folder Structure Mapping to Word Document

This project generates a Word document representing the structure of a network folder. The document includes the hierarchy of folders and files with proper indentation and color-coding based on file types. It's perfect for creating a visual map of a network's directory structure.

![Demo](./demo.gif)

## Features

- Generates a Word document with folder and file hierarchy.
- Color-codes files based on their extensions.
- Adjusts the font size of file names based on their depth in the folder structure.

## Requirements

- Python 3.x
- `python-docx` library
- `tqdm` library

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/MathFreedom/Folder-Structure-Mapping-to-Word-Document.git
    cd network-folder-structure-mapping
    ```

2. Install the required packages:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

### Default Paths
The script can also be run with default input and output paths:

```bash
python folder_structure_to_word.py

### Command Line

To run the script from the command line, use:
```bash
python folder_structure_to_word.py <folder_path> <output_doc_path>
