import os, sys #Standard Python Libraries
from docxtpl import DocxTemplate

# Change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate('hmke.docx')
context = {'iktatószám': '1-cég1-2023'}

doc.render(context)
doc.save('hmke_rendered.docx')