#!/usr/bin/env python3
"""
run_level_rsid_analysis.py

Analyzes a Word document's XML to identify and categorize revision markers
at the run (<w:r>) level, focusing specifically on w:rsidRPr.

This version incorporates advanced logic:
1.  It groups contiguous runs with the same w:rsidRPr to treat them as a
    single logical change.
2.  It intelligently finds the last preceding run containing text (<w:t>) to
    establish a more accurate "before" state for comparison, leading to
    a more meticulous and robust analysis.
"""
import sys
from lxml import etree as ET
from collections import defaultdict

# Namespace for WordprocessingML
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

        if current_val is not None and base_val is None:
            # Property was added
            val = current_val.get(attr_key, 'N/A') if attr_key else ''
            reasons.append(reason_template.format(val).strip())
        elif current_val is not None and base_val is not None:
            # Property might have changed
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

    # Start with the paragraph's default run properties, if they exist
    if paragraph is not None and paragraph.tag == f'{{{NS["w"]}}}p':
        p_rpr = paragraph.find('w:pPr/w:rPr', namespaces=NS)
        base_props.update(get_properties_dict(p_rpr))

    # Search backwards from the current run for the last one with text
    current = run_element.getprevious()
    while current is not None:
        if current.tag == f'{{{NS["w"]}}}r' and current.find('w:t', namespaces=NS) is not None:
            # Found the last relevant run. Its properties are our base.
            prev_rpr = current.find('w:rPr', namespaces=NS)
            base_props.update(get_properties_dict(prev_rpr))
            return base_props # Exit once found
        current = current.getprevious()

    return base_props

def analyze_document_runs(document_xml_path):
    """
    Parses document.xml, groups contiguous runs with the same w:rsidRPr,
    and analyzes them to determine the cause of the revision.
    """
    try:
        doc_tree = ET.parse(document_xml_path)
    except (ET.XMLSyntaxError, FileNotFoundError) as e:
        print(f"  Error processing document file: {e}")
        return None, 0

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

    analysis_map = defaultdict(list)
    total_rsidrpr_runs = 0

    # Iterate through paragraphs to handle run grouping correctly
    for p_element in doc_tree.xpath("//w:p", namespaces=NS):
        runs = p_element.xpath("./w:r", namespaces=NS)
        i = 0
        while i < len(runs):
            run = runs[i]
            rsid_rpr = run.get(f'{{{NS["w"]}}}rsidRPr')

            if not rsid_rpr:
                i += 1
                continue

            total_rsidrpr_runs += 1
            # --- Grouping Logic ---
            group = [run]
            j = i + 1
            while j < len(runs) and runs[j].get(f'{{{NS["w"]}}}rsidRPr') == rsid_rpr:
                group.append(runs[j])
                total_rsidrpr_runs += 1
                j += 1
            
            # --- Analysis of the Group ---
            first_run_in_group = group[0]
            base_props = find_effective_prior_properties(first_run_in_group)
            
            current_rpr = first_run_in_group.find('w:rPr', namespaces=NS)
            current_props = get_properties_dict(current_rpr)
            
            reasons = diff_properties(base_props, current_props, prop_definitions)
            
            # Combine text from all runs in the group
            full_text = "".join(
                "".join(r.xpath('.//w:t/text()', namespaces=NS)) for r in group
            ).strip()

            if not reasons:
                reasons.append("No property change detected") #(likely a split run or style refresh)

            analysis_map[rsid_rpr].append({
                "text": full_text or "[No Text in Run(s)]",
                "reasons": reasons
            })
            
            # Advance the main loop past the processed group
            i = j

    return analysis_map, total_rsidrpr_runs

def main():
    if len(sys.argv) != 3:
        print("Usage: python run_level_rsid_analysis.py <document.xml> <output.txt>")
        sys.exit(1)
        
    document_xml_path, output_path = sys.argv[1:3]
    
    analysis_results, run_count = analyze_document_runs(document_xml_path)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("--- Run-Level w:rsidRPr Analysis (Grouped) ---\n\n")
        f.write(f"Found and analyzed {run_count} individual <w:r> elements with a w:rsidRPr attribute.\n")
        f.write("NOTE: Contiguous runs with the same rsidRPr are grouped into a single entry below.\n\n")
        
        if not analysis_results:
            f.write("No runs with a w:rsidRPr attribute were found.\n")
            return

        for rsid in sorted(analysis_results.keys()):
            f.write(f"RSID: {rsid}\n")
            f.write("-" * 25 + "\n")
            for entry in analysis_results[rsid]:
                f.write(f"  Combined Text: \"{entry['text']}\"\n")
                f.write(f"  Identified Change(s):\n")
                for reason in entry['reasons']:
                    f.write(f"    - {reason}\n")
            f.write("\n")

    print(f"  Run-level analysis results written to '{output_path}'.")

if __name__ == "__main__":
    main()