import random
from docx import Document
from lxml import etree

# Open the Word document
doc = Document('test.docx')

# Create an XML element for the HTML document
html = etree.Element('html')
head = etree.SubElement(html, 'head')
body = etree.SubElement(html, 'body')

# Iterate through each paragraph in the document
previous_was_paragraph = False
for paragraph in doc.paragraphs:
    if paragraph.text.strip():
        # Create a new HTML paragraph element and add the text
        p = etree.SubElement(body, 'p')
        
        # Add a self-closing <a> tag at the start of the paragraph
        a = etree.Element('a')
        a.set('class', 'brl-location')
        a.set('id', str(random.randint(100000000, 999999999)))
        p.append(a)

        p.text = paragraph.text.strip()

# Serialize the HTML element to a string
html_string = etree.tostring(html, encoding='unicode', pretty_print=True)

# Save the HTML string to a file
with open('test.html', 'w') as file:
    file.write(html_string)
