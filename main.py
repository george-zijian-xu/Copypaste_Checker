# backend/src/main.py
"""
Main entry point for the Word Formatter pipeline.
Orchestrates the unzipping and analysis of a .docx file.
"""

import sys
import os
import subprocess
from config import Paths, Messages

def run_command(command, description):
    """Run a subprocess command and handle errors."""
    print(f"\n {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f" {description} completed successfully.")
        if result.stdout:
            print(result.stdout.strip())
        return True
    except subprocess.CalledProcessError as e:
        print(f" Error during {description}:")
        print(f"   Command: {command}")
        print(f"   Error: {e.stderr}")
        return False

def main():
    """Main pipeline execution."""
    if len(sys.argv) != 2:
        print(Messages.USAGE_MAIN)
        sys.exit(1)

    docx_path = sys.argv[1]

    if not os.path.isfile(docx_path):
        print(Messages.FILE_NOT_FOUND.format(docx_path))
        sys.exit(1)

    print(f" Starting Word Analysis Pipeline for: {docx_path}")

    # --- Directory Setup ---
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    output_dir = os.path.join(Paths.OUTPUT_DIR, base_name)
    unzipped_dir = os.path.join(output_dir, "unzipped")
    
    os.makedirs(output_dir, exist_ok=True)
    print(f" Created main output directory: {output_dir}")

    # --- File Paths ---
    document_xml_path = os.path.join(unzipped_dir, "word", Paths.DOCUMENT_XML_NAME)
    settings_xml_path = os.path.join(unzipped_dir, "word", Paths.SETTINGS_XML_NAME)
    analysis_output_path = os.path.join(output_dir, Paths.ANALYSIS_OUTPUT_NAME)
    p_tags_output_path = os.path.join(output_dir, Paths.P_TAGS_OUTPUT_NAME)
    run_analysis_output_path = os.path.join(output_dir, Paths.RUN_ANALYSIS_OUTPUT_NAME) # New path

    # --- Pipeline Steps ---

    # Step 1: Unzip the .docx file
    unzip_command = f"python unzip.py \"{docx_path}\" \"{unzipped_dir}\""
    if not run_command(unzip_command, f"Unzipping .docx to '{unzipped_dir}'"):
        sys.exit("❌ Pipeline failed at unzip step.")

    # Check if the required XML files exist before proceeding
    if not os.path.exists(document_xml_path) or not os.path.exists(settings_xml_path):
        print(f"❌ Error: Could not find '{Paths.DOCUMENT_XML_NAME}' or '{Paths.SETTINGS_XML_NAME}'.")
        print("   Please ensure the .docx file is valid and contains these files.")
        sys.exit("❌ Pipeline halted.")

    # Step 2: Perform Paragraph-level RSID Analysis
    analysis_command = (
        f"python rsid_analysis.py \"{document_xml_path}\" "
        f"\"{settings_xml_path}\" \"{analysis_output_path}\""
    )
    if not run_command(analysis_command, "Analyzing RSIDs based on settings.xml"):
        print("  Continuing pipeline despite analysis failure.")

    # Step 3: Extract all <w:p> tags
    p_tags_command = (
        f"python extract_p_tags.py \"{document_xml_path}\" \"{p_tags_output_path}\""
    )
    if not run_command(p_tags_command, "Extracting all paragraph tags"):
        print("  Continuing pipeline despite p-tag extraction failure.")

    # Step 4: Perform Run-level rsidR Analysis (New Step)
    run_analysis_command = (
        f"python run_level_rsid_analysis.py \"{document_xml_path}\" \"{run_analysis_output_path}\""
    )
    if not run_command(run_analysis_command, "Analyzing run-level rsidR properties"):
        print("  Continuing pipeline despite run-level analysis failure.")

    print(f"\n {Messages.SUCCESS_PIPELINE}")
    print(f" Paragraph-level RSID analysis results: {analysis_output_path}")
    print(f" Paragraph tag data: {p_tags_output_path}")
    print(f" Run-level rsidR analysis results: {run_analysis_output_path}")

if __name__ == "__main__":
    main()