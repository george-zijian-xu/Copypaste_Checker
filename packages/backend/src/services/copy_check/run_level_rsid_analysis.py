#!/usr/bin/env python3
"""
run_level_rsid_analysis.py

Analyzes a Word document's XML to identify and categorize revision markers
at the run (<w:r>) level, focusing specifically on w:rsidRPr.

This version has been refactored to:
1.  Output a JSON object containing a list of "highlights" with start/end
    character offsets relative to the full text.
2.  Use a centralized text extractor to ensure offsets are accurate.
3.  Incorporate advanced logic to group contiguous runs and differentiate
    font face changes from font hint changes.
"""
import sys
import json
from lxml import etree as ET
from collections import defaultdict

# Namespace for WordprocessingML is now defined locally
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_properties_dict(rpr_element):
    """Converts a <w:rPr> element into a dictionary of its properties."""
    if rpr_element is None:
        return {}
    props = {child.tag: dict(child.attrib) for child in rpr_element}
    return props

def diff_properties(base_props, current_props, prop_definitions):
    """
    Compares two dictionaries of run properties and returns human-readable
    descriptions of what was added, changed, or removed.
    """
    reasons = []
    for key, (reason_template, attr_key) in prop_definitions.items():
        base_val = base_props.get(key)
        current_val = current_props.get(key)

        # Special handling for <w:rFonts> to distinguish between a face and hint change.
        if key.endswith('}rFonts') and current_val is not None:
            base_ascii = (base_val or {}).get(f'{{{NS["w"]}}}ascii')
            base_hint  = (base_val or {}).get(f'{{{NS["w"]}}}hint')
            curr_ascii = current_val.get(f'{{{NS["w"]}}}ascii')
            curr_hint  = current_val.get(f'{{{NS["w"]}}}hint')

            if base_val is None:
                if curr_ascii:
                    reasons.append(f"Formatting: Font Change (Face: {curr_ascii})")
                elif curr_hint:
                    reasons.append(f"Formatting: Font Hint Change (Hint: {curr_hint})")
                else:
                    reasons.append("Formatting: Font Change (Face: N/A)")
            else:
                if base_ascii != curr_ascii:
                    reasons.append(f"Formatting: Font Change (Face: {curr_ascii or 'N/A'})")
                if base_hint != curr_hint:
                    reasons.append(f"Formatting: Font Hint Change (Hint: {curr_hint or 'N/A'})")
            continue

        if current_val is not None and base_val is None:
            val = current_val.get(attr_key, 'N/A') if attr_key else ''
            reasons.append(reason_template.format(val).strip())
        elif current_val is not None and base_val is not None:
            if attr_key and base_val.get(attr_key) != current_val.get(attr_key):
                val = current_val.get(attr_key, 'N/A')
                reasons.append(reason_template.format(val).strip())
    return reasons

def find_effective_prior_properties(run_element):
    """
    Finds the most relevant preceding run properties by searching backwards
    for the first run that contains a <w:t> element.
    """
    base_props = {}
    paragraph = run_element.getparent()
    if paragraph is not None and paragraph.tag == f'{{{NS["w"]}}}p':
        p_rpr = paragraph.find('w:pPr/w:rPr', namespaces=NS)
        base_props.update(get_properties_dict(p_rpr))

    current = run_element.getprevious()
    while current is not None:
        if current.tag == f'{{{NS["w"]}}}r' and current.find('w:t', namespaces=NS) is not None:
            prev_rpr = current.find('w:rPr', namespaces=NS)
            base_props.update(get_properties_dict(prev_rpr))
            return base_props
        current = current.getprevious()
    return base_props

def analyze_document_runs(document_xml_path):
    """
    Parses document.xml to extract plain text and identify revision markers.
    
    This function is self-contained: it generates the text and a list of
    highlights with character offsets in a single pass, ensuring that the
    offsets are always correct relative to the generated text.

    Returns:
        A tuple: (full_text: str, highlights: list)
    """
    try:
        doc_tree = ET.parse(document_xml_path)
    except (ET.XMLSyntaxError, FileNotFoundError) as e:
        print(f"  Error processing document file: {e}", file=sys.stderr)
        return "", []

    prop_definitions = {
        f'{{{NS["w"]}}}b': ("Formatting: Bold", None),
        f'{{{NS["w"]}}}i': ("Formatting: Italic", None),
        f'{{{NS["w"]}}}u': ("Formatting: Underline (Style: {0})", f'{{{NS["w"]}}}val'),
        f'{{{NS["w"]}}}color': ("Formatting: Color (Value: {0})", f'{{{NS["w"]}}}val'),
        f'{{{NS["w"]}}}highlight': ("Formatting: Highlight (Color: {0})", f'{{{NS["w"]}}}val'),
        f'{{{NS["w"]}}}strike': ("Formatting: Strikethrough", None),
        f'{{{NS["w"]}}}sz': ("Formatting: Font Size (Value: {0})", f'{{{NS["w"]}}}val'),
        f'{{{NS["w"]}}}rFonts': ("Formatting: Font Change (Face: {0})", f'{{{NS["w"]}}}ascii'), 
        f'{{{NS["w"]}}}proofErr': ("Proofing: Spelling Error Flag (Type: {0})", f'{{{NS["w"]}}}type'),
        f'{{{NS["w"]}}}gramErr': ("Proofing: Grammar Error Flag", None),
        f'{{{NS["w"]}}}lang': ("Property: Language Change (Lang: {0})", f'{{{NS["w"]}}}val'),
        f'{{{NS["w"]}}}rStyle': ("Property: Character Style Applied (Style: {0})", f'{{{NS["w"]}}}val'),
    }

    highlights = []
    text_parts = []

    # Use XPath to process only paragraphs outside of tables
    paragraphs = doc_tree.xpath(".//w:p[not(ancestor::w:tbl)]", namespaces=NS)

    for p_element in paragraphs:
        runs = p_element.xpath("./w:r", namespaces=NS)
        i = 0
        while i < len(runs):
            run = runs[i]
            
            # --- Text Extraction and Offset Calculation ---
            # Calculate current position before adding new text
            current_pos = len("".join(text_parts))
            run_text_parts = [t.text or "" for t in run.xpath('.//w:t', namespaces=NS)]
            run_text = "".join(run_text_parts)

            rsid_rpr = run.get(f'{{{NS["w"]}}}rsidRPr')

            if not rsid_rpr:
                text_parts.append(run_text)
                i += 1
                continue

            # --- Grouping Logic ---
            group = [run]
            group_text_parts = run_text_parts
            j = i + 1
            while j < len(runs) and runs[j].get(f'{{{NS["w"]}}}rsidRPr') == rsid_rpr:
                next_run = runs[j]
                next_text_parts = [t.text or "" for t in next_run.xpath('.//w:t', namespaces=NS)]
                group_text_parts.extend(next_text_parts)
                group.append(next_run)
                j += 1
            
            full_group_text = "".join(group_text_parts)
            text_parts.append(full_group_text)
            
            # --- Analysis of the Group ---
            first_run_in_group = group[0]
            base_props = find_effective_prior_properties(first_run_in_group)
            current_rpr = first_run_in_group.find('w:rPr', namespaces=NS)
            current_props = get_properties_dict(current_rpr)
            reasons = diff_properties(base_props, current_props, prop_definitions)
            
            if not reasons and full_group_text:
                reasons.append("No property change detected")

            if reasons:
                start = current_pos
                end = start + len(full_group_text)
                for reason in reasons:
                    highlights.append({
                        "start": start,
                        "end": end,
                        "category": reason,
                        "rsid": rsid_rpr,
                    })
            i = j
        
        # Add a newline between paragraphs
        text_parts.append("\n")

    # Join all parts, but remove the final extraneous newline
    full_text = "".join(text_parts).rstrip("\n")

    return full_text, highlights

def main():
    if len(sys.argv) != 3:
        print("Usage: python run_level_rsid_analysis.py <document.xml> <output.json>", file=sys.stderr)
        sys.exit(1)
        
    document_xml_path, output_path = sys.argv[1:3]
    
    full_text, analysis_results = analyze_document_runs(document_xml_path)
    
    output_data = {
        "sourceText": full_text,
        "highlights": analysis_results
    }
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2)

    print(f"  Run-level analysis results written to '{output_path}'.")

if __name__ == "__main__":
    main()