import sys
from lxml import etree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def extract_text(document_xml_path: str) -> str:
    """
    Extracts all paragraph text from a document.xml file, skipping tables.
    Joins paragraphs with a single newline character.
    """
    try:
        doc_tree = ET.parse(document_xml_path)
    except (ET.XMLSyntaxError, FileNotFoundError) as e:
        print(f"Error processing document file for text extraction: {e}", file=sys.stderr)
        return ""

    # XPath to find all paragraphs <w:p> that are NOT inside a table <w:tbl>
    xpath = ".//w:p[not(ancestor::w:tbl)]"
    paragraphs = doc_tree.xpath(xpath, namespaces=NS)
    
    paragraph_texts = []
    for p in paragraphs:
        # For each paragraph, get all its text runs <w:t> and join them
        text = "".join(t for t in p.xpath(".//w:t/text()", namespaces=NS))
        paragraph_texts.append(text)
        
    return "\n".join(paragraph_texts)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python docx_to_text.py <document.xml>", file=sys.stderr)
        sys.exit(1)
    
    document_xml_path = sys.argv[1]
    full_text = extract_text(document_xml_path)
    print(full_text) 