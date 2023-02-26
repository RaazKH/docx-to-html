import os
from docx import Document
from lxml import etree
import random

# Set the input and output directory paths
input_dir = 'word'
output_dir = 'html'

# Create the output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

# Iterate through each file in the input directory
for filename in os.listdir(input_dir):
    # Check if the file is a Word document
    if filename.endswith('.docx'):
        # Open the Word document
        doc = Document(os.path.join(input_dir, filename))

        # Create an XML element for the HTML document
        html = etree.Element('html')
        body = etree.SubElement(html, 'body')

        # Iterate through each paragraph in the document
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
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

        # Save the HTML string to a file with the same name as the Word file
        output_filename = os.path.splitext(filename)[0] + '.html'
        with open(os.path.join(output_dir, output_filename), 'w') as file:
            file.write(html_string)
