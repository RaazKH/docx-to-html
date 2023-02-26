from docx import Document
from lxml import etree
import os
import random

# Set the paths for the input and output directories
input_dir = 'word'
output_dir = 'html'

# Create the output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Iterate through each file in the input directory
for filename in os.listdir(input_dir):
    # Check if the file is a Word document
    if filename.endswith('.docx') and not filename.startswith('~$'):
        # Open the Word document
        doc_path = os.path.join(input_dir, filename)
        doc = Document(doc_path)

        # Create an XML element for the HTML document
        html = etree.Element('html')
        body = etree.SubElement(html, 'body')

        # Iterate through each paragraph in the document
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                # Check if the paragraph starts with "This document" or "Last Modified"
                if paragraph.text.strip().startswith('This document') or paragraph.text.strip().startswith('Last Modified'):
                    # Skip the paragraph
                    continue

                # Create a new HTML paragraph element and add the text
                p = etree.Element('p')

                # Add a self-closing <a> tag at the start of the paragraph
                a_id = str(random.randint(100000000, 999999999))
                a = etree.Element('a', {'class': 'brl-location', 'id': a_id})
                a.tail = paragraph.text.strip()
                p.append(a)

                body.append(p)

        # Serialize the HTML element to a string
        html_string = etree.tostring(html, encoding='unicode', pretty_print=True)

        # Create the output file path
        html_path = os.path.join(output_dir, os.path.splitext(filename)[0] + '.html')

        # Save the HTML string to a file
        with open(html_path, 'w') as file:
            file.write(html_string)
