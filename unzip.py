#!/usr/bin/env python3
"""
A simple utility to unzip a .docx file to a specified directory.
"""
import sys
import os
import zipfile

def unzip_docx(docx_path: str, output_dir: str):
    """
    Unzips the given .docx file into the specified output directory.
    Returns True on success, False on failure.
    """
    if not os.path.isfile(docx_path):
        print(f" Error: Input file not found at '{docx_path}'")
        return False
    
    try:
        # Ensure the output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        with zipfile.ZipFile(docx_path, 'r') as zf:
            zf.extractall(output_dir)
        # No success message here to keep the utility quiet
        return True
    except zipfile.BadZipFile:
        print(f" Error: The file '{docx_path}' is not a valid zip archive.")
        return False
    except Exception as e:
        print(f" An unexpected error occurred during unzipping: {e}")
        return False

if __name__ == "__main__":
    # This part is for direct command-line execution for testing
    if len(sys.argv) != 3:
        print("Usage: python unzip.py <path_to_docx> <output_directory>")
        sys.exit(1)

    if unzip_docx(sys.argv[1], sys.argv[2]):
        print(f" Successfully unzipped '{sys.argv[1]}' to '{sys.argv[2]}'")
    else:
        print(" Unzipping failed.")