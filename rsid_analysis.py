#!/usr/bin/env python3
"""
rsid_analysis.py

Analyzes a Word document's XML to extract and combine text based on revision
markers (rsid) defined in settings.xml. It correctly handles RSIDs at the
paragraph level.
"""
import sys
from collections import defaultdict
from lxml import etree as ET

# Namespace for WordprocessingML
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_rsids_from_settings(settings_xml_path):
    """Parses settings.xml to extract all unique RSID values."""
    try:
        tree = ET.parse(settings_xml_path)
        rsid_nodes = tree.findall('.//w:rsid', namespaces=NS)
        rsids = {node.get(f'{{{NS["w"]}}}val') for node in rsid_nodes}
        print(f" Found {len(rsids)} unique RSIDs in '{settings_xml_path}'.")
        return rsids
    except (ET.XMLSyntaxError, FileNotFoundError) as e:
        print(f" Error processing settings file: {e}")
        return set()

def analyze_document_revisions(document_xml_path, rsids):
    """
    Searches document.xml for text within paragraphs corresponding to the
    given RSIDs and combines it.
    """
    rsid_text_map = defaultdict(list)
    
    try:
        doc_tree = ET.parse(document_xml_path)
        
        # Find all paragraphs in the document that have a w:rsidR attribute.
        # This is much more efficient than running a separate query for every single RSID.
        query = ".//w:p[@w:rsidR]"
        
        paragraphs_with_rsid = doc_tree.xpath(query, namespaces=NS)

        for p_element in paragraphs_with_rsid:
            # Get the rsidR value from this paragraph
            p_rsid = p_element.get(f'{{{NS["w"]}}}rsidR')
            
            # If this paragraph's RSID is one we're looking for...
            if p_rsid in rsids:
                # Find all text nodes within this paragraph and append their text
                text_nodes = p_element.xpath('.//w:t', namespaces=NS)
                for t_node in text_nodes:
                    if t_node.text:
                        rsid_text_map[p_rsid].append(t_node.text)

        # Join the collected text fragments for each RSID
        return {rsid: "".join(texts) for rsid, texts in rsid_text_map.items()}

    except (ET.XMLSyntaxError, FileNotFoundError) as e:
        print(f"  Error processing document file: {e}")
        return {}

def main():
    if len(sys.argv) != 4:
        # Use a generic message for standalone execution
        print("Usage: python rsid_analysis.py <document.xml> <settings.xml> <output.txt>")
        sys.exit(1)
        
    document_xml_path, settings_xml_path, output_path = sys.argv[1:4]
    
    all_rsids = get_rsids_from_settings(settings_xml_path)
    if not all_rsids:
        print("  No RSIDs found in settings.xml. Cannot perform analysis.")
        return

    rsid_text_map = analyze_document_revisions(document_xml_path, all_rsids)
    
    found_rsids = set(rsid_text_map.keys())
    missing_rsids = all_rsids - found_rsids
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("--- RSID Text Analysis ---\n\n")
        
        if rsid_text_map:
            f.write("--- Found Revisions ---\n")
            for rsid in sorted(rsid_text_map.keys()):
                text = rsid_text_map[rsid]
                f.write(f"RSID: {rsid}\n")
                f.write(f"Text: {text.strip()}\n")
                f.write("-" * 25 + "\n")
        else:
            f.write("No text associated with the provided RSIDs was found in the document.\n")

        if missing_rsids:
            f.write("\n\n--- Not Found in Document ---\n")
            f.write("The following RSIDs from settings.xml had no corresponding text in document.xml:\n")
            for rsid in sorted(missing_rsids):
                f.write(f"- {rsid}\n")

    print(f"  Analysis results written to '{output_path}'.")
    print(f"   - Found text for {len(found_rsids)} RSIDs.")
    print(f"   - Not found: {len(missing_rsids)} RSIDs.")

if __name__ == "__main__":
    main()