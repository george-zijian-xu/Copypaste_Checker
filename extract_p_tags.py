#!/usr/bin/env python3
"""
extract_p_tags.py

A utility to extract the full opening tag of every <w:p> element from a
document.xml file. This version correctly reconstructs the tag and removes
unnecessary namespace (xmlns) declarations from the output.
"""
import sys
from lxml import etree as ET

def extract_paragraph_tags(document_xml_path, output_path):
    """
    Parses document.xml and writes the clean opening <w:p> tags to a file.
    """
    try:
        # Use iterparse for memory efficiency. We listen for 'end' events.
        context = ET.iterparse(document_xml_path, events=('end',), tag='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

        with open(output_path, 'w', encoding='utf-8') as f:
            for event, elem in context:
                # --- Tag Name Reconstruction ---
                # The elem.tag is in the format {namespace}tagname. We need to find the prefix.
                prefix = None
                # Find the prefix (like 'w') for the element's namespace URI.
                for p, ns_uri in elem.nsmap.items():
                    if ns_uri == ET.QName(elem).namespace:
                        prefix = p
                        break
                
                # Construct the prefixed tag name (e.g., 'w:p')
                tag_name = f"{prefix}:{ET.QName(elem).localname}" if prefix else ET.QName(elem).localname


                # --- Attribute Reconstruction (stripping xmlns) ---
                attrs_list = []
                for attr_name, attr_value in elem.attrib.items():
                    # Get the prefix for the attribute's namespace
                    attr_prefix = None
                    q_attr = ET.QName(attr_name)
                    # Find the prefix for the attribute's namespace, if it has one
                    for p, ns_uri in elem.nsmap.items():
                        if ns_uri == q_attr.namespace:
                            attr_prefix = p
                            break
                    
                    # Construct the prefixed attribute name (e.g., 'w14:paraId')
                    full_attr_name = f"{attr_prefix}:{q_attr.localname}" if attr_prefix else q_attr.localname
                    attrs_list.append(f'{full_attr_name}="{attr_value}"')

                attrs_str = ' '.join(attrs_list)
                
                # --- Final Tag Assembly ---
                full_tag = f"<{tag_name} {attrs_str}>"
                f.write(full_tag + "\n")

                # Clear the element and its ancestors to free memory
                elem.clear()
                while elem.getprevious() is not None:
                    del elem.getparent()[0]

        print(f" Successfully extracted all <w:p> tags to '{output_path}'.")
        return True

    except (ET.XMLSyntaxError, FileNotFoundError) as e:
        print(f" Error processing document file: {e}")
        return False
    except Exception as e:
        print(f" An unexpected error occurred: {e}")
        return False


def main():
    if len(sys.argv) != 3:
        print("Usage: python extract_p_tags.py <document.xml> <output.txt>")
        sys.exit(1)
    
    document_xml_path, output_path = sys.argv[1:3]
    extract_paragraph_tags(document_xml_path, output_path)

if __name__ == "__main__":
    main()