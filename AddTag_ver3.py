import os
import docx
from docx import Document
import datetime
import shutil
import pandas as pd

def add_tag_after_heading(doc, heading, tag, filename):
    tag_added = False
    for i in range(len(doc.paragraphs)):
        paragraph = doc.paragraphs[i]
        if 'Heading' in paragraph.style.name and paragraph.text == heading:
            # Add a new paragraph with the tag after the heading
            new_paragraph = doc.add_paragraph(tag)
            # Move the new paragraph to be after the heading
            doc.element.body.insert(i+1, new_paragraph._element)
            tag_added = True
    if tag_added:
        if filename not in tag_added_files:
            tag_added_files[filename] = []
        tag_added_files[filename].append(tag)

def export_to_excel(tag_added_files):
    writer = pd.ExcelWriter('Tag_Added_Report.xlsx', engine='openpyxl')
    for filename in os.listdir(input_folder):
        data = {"Heading": headings, "Tag Added": []}
        for heading in headings:
            if filename in tag_added_files and heading in tag_added_files[filename]:
                data["Tag Added"].append("Yes")
            else:
                data["Tag Added"].append("No")
        df = pd.DataFrame(data)
        df.to_excel(writer, sheet_name=filename, index=False)
    writer.save()

input_folder = 'path/to/Input'
output_folder = 'path/to/Output'
headings = ['ABCD', 'EFGH', 'IJKL']
tags = ['<ABCD></>', '<EFGH></>', '<IJKL></>']
tag_added_files = {}

# Generate a new folder with the current timestamp
current_timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
output_folder = os.path.join(output_folder, f"{current_timestamp}_AddTag")
os.mkdir(output_folder)

# Process each Word document in the input folder
for filename in os.listdir(input_folder):
    if filename.endswith('.docx'):
        file_path = os.path.join(input_folder, filename)
        doc = Document(file_path)
        # Add the appropriate tag after each heading
        for heading, tag in zip(headings, tags):
            add_tag_after_heading(doc, heading, tag, filename)
        # Save the updated document to the output folder
        output_file_path = os.path.join(output_folder, filename[:-5] + '_AddTag.docx')
        doc.save(output_file_path)

# Export the results to Excel
export_to_excel(tag_added_files)
