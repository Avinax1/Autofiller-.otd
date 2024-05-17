import zipfile
from lxml import etree
from docx import Document


def read_docx(file_path):
    doc = Document(file_path)
    return [p.text for p in doc.paragraphs if p.text]


def replace_placeholders(odt_path, replacements):
    ns = {
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
    }

    # Read the content of the ODT file
    with zipfile.ZipFile(odt_path, 'r') as odt_file:
        content_xml = odt_file.read('content.xml')

    # Parse the content XML
    root = etree.fromstring(content_xml)
    paragraphs = root.findall('.//text:p', namespaces=ns)

    # Replace placeholders in the paragraphs
    for paragraph in paragraphs:
        text_content = ''.join(paragraph.itertext())
        for key, value in replacements.items():
            if key in text_content:
                print(f"Replacing '{key}' with '{value}'")
                text_content = text_content.replace(key, value)
                # Clear current paragraph content
                for elem in list(paragraph):
                    paragraph.remove(elem)
                # Set new text content
                paragraph.text = text_content

    # Save the modified content back to the ODT file
    with zipfile.ZipFile(odt_path, 'a') as odt_file:
        with odt_file.open('content.xml', 'w') as content_file:
            content_file.write(etree.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8'))


def main(docx_path, odt_path):
    # Read data from .docx file
    paragraphs = read_docx(docx_path)

    # Prepare replacements dictionary
    replacements = {
        '[b]': paragraphs[1] if len(paragraphs) > 1 else '',
        '[g]': paragraphs[4] if len(paragraphs) > 4 else '',
        '[h]': paragraphs[2] if len(paragraphs) > 2 else '',
        '[k]': paragraphs[5] if len(paragraphs) > 5 else '',
        '[a]': paragraphs[3] if len(paragraphs) > 3 else '',
        '[c]': paragraphs[6] if len(paragraphs) > 6 else '',
    }

    # Replace placeholders in .odt file
    replace_placeholders(odt_path, replacements)


if __name__ == "__main__":
    docx_path = ['input.docx']
    odt_path = ['file.odt']
    main(docx_path, odt_path)
