from docx import Document
from lxml import etree

# Open the Word document
doc = Document('test.docx')

# Create an XML element for the HTML document
html = etree.Element('html')
body = etree.SubElement(html, 'body')

# Iterate through each paragraph in the document
previous_was_paragraph = False
for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    if text:
        # Create a new HTML paragraph element and add the text
        p = etree.SubElement(body, 'p')
        p.text = text

# Serialize the HTML element to a string
html_string = etree.tostring(html, encoding='unicode', pretty_print=True)

# Save the HTML string to a file
with open('test.html', 'w') as file:
    file.write(html_string)
