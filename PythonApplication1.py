

#pip install python-docx
#start from here if python-docx is installed
from docx import Document
#indentation lib
from docx.shared import Inches
#font lib
from docx.shared import Pt
#open the document
document=Document('C:\Program Files\Python Programs\cover-letter\cover-letter-python\PythonApplication1\PythonApplication1\\test.docx')
newword=input('Enter Company Name:')
#getting the style object of the doc
style = document.styles['Normal']
#modifying the style
font = style.font
font.name = 'Calibri'
font.size = Pt(12)
for paragraph in document.paragraphs:
    if "your Company" in paragraph.text:
        print (paragraph.text)
        #replacing your word or sentence with another string
        paragraph.text=paragraph.text.replace("your Company", newword)
        #paragraph.paragraph_format.line_spacing = Inches(0.5)
        print('=======================')
        print(paragraph.text)
        #assigning the created style to the edited paragraph
        paragraph.style = document.styles['Normal']
#save changed document

document.save('C:\Program Files\Python Programs\cover-letter\cover-letter-python\PythonApplication1\PythonApplication1\\test1.docx')
