import os
import datetime
import pandas as pd
from docx import Document

def extract_headings_from_docx(docx_path):
    document = Document(docx_path)
    headings = []
    for paragraph in document.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            headings.append(paragraph.text)
    return headings

def save_headings_to_excel(headings_dict, output_dir):
    timestamp = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
    output_file = os.path.join(output_dir, f'DetectSection_{timestamp}.xlsx')
    
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for filename, headings in headings_dict.items():
        df = pd.DataFrame(headings, columns=['Headings'])
        df.to_excel(writer, sheet_name=filename, index=False)

    writer.save()

def main(input_dir, output_dir):
    headings_dict = {}

    for filename in os.listdir(input_dir):
        if filename.endswith('.docx'):
            docx_path = os.path.join(input_dir, filename)
            headings = extract_headings_from_docx(docx_path)
            headings_dict[filename] = headings

    save_headings_to_excel(headings_dict, output_dir)

if __name__ == '__main__':
    input_dir = '/path/to/your/input/folder'  # Change to your input folder
    output_dir = '/path/to/your/output/folder'  # Change to your output folder
    main(input_dir, output_dir)
