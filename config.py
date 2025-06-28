"""
Configuration file for the Word Formatter pipeline.
"""

class Paths:
    # Top-level directories
    OUTPUT_DIR = "output"
    
    # Key file names
    DOCUMENT_XML_NAME = "document.xml"
    SETTINGS_XML_NAME = "settings.xml"
    ANALYSIS_OUTPUT_NAME = "rsid_analysis.txt"
    P_TAGS_OUTPUT_NAME = "p_tags_output.txt"
    RUN_ANALYSIS_OUTPUT_NAME = "run_level_rsid_analysis.txt" # New entry

class Messages:
    FILE_NOT_FOUND = "Error: File not found: {}"
    USAGE_MAIN = "Usage: python main.py <path_to_docx>"
    USAGE_UNZIP = "Usage: python unzip.py <path_to_docx> <output_directory>"
    USAGE_ANALYSIS = "Usage: python rsid_analysis.py <document.xml> <settings.xml> <output.txt>"
    
    SUCCESS_PIPELINE = "Pipeline completed successfully!"
    SUCCESS_UNZIP = "Successfully unzipped '{}' to '{}'."