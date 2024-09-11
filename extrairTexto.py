import os
import pdfplumber

curriculo = os.listdir()

pdf = pdfplumber.open('estadia\CV Dev - Adriel (1).pdf')

page = pdf.pages[0]
text = page.extract_text()
print(text.split('\n'))
nome = text.split('\n')[0]

print('\nnome = '+ nome)