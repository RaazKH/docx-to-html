from docx import Document
from lxml import etree
import random

# Open the Word document
doc = Document('test.docx')

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

# Save the HTML string to a file
with open('test.html', 'w') as file:
    file.write(html_string)
