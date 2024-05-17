from docx import Document

def read_data_after_colon(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            data = file.read()
            lines = data.splitlines()
            data_after_colon = []
            for i, line in enumerate(lines):
                if ':' in line:
                    key, value = line.split(':', 1)
                    data_after_colon.append(value.strip())
                    # Sprawdzamy, czy klucz zawiera słowa kluczowe sugerujące "Firmę" lub "Adres"
                    if any(keyword in key for keyword in ["Firma", "Adres"]) or key.strip().isdigit():
                        if i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            data_after_colon.append(next_line)
                            print(f"Key: '{key}', Value: '{value}', Next line: '{next_line}'")
            return data_after_colon
    except FileNotFoundError:
        return None

# Example usage:
input_filename = 'mail.txt'
data_to_insert = read_data_after_colon(input_filename)
if data_to_insert:
    print("Data to insert:")
    print(data_to_insert)
else:
    print(f"File '{input_filename}' not found.")



def insert_data_into_word(data, docx_filename):
    doc = Document()
    for line in data:
        doc.add_paragraph(line.strip())
    doc.save(docx_filename)
    print(f"Data has been placed in the file {docx_filename}")

# Example usage:
input_filename = 'mail.txt'
output_docx_filename = 'output.docx'
data_to_insert = read_data_after_colon(input_filename)
if data_to_insert:
    insert_data_into_word(data_to_insert, output_docx_filename)
else:
    print(f"File '{input_filename}' not found.")
