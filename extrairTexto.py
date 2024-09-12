import pdfplumber
pdf = pdfplumber.open('estadia\TICKET 1.pdf')
page = pdf.pages[0]
text = page.extract_text()
# dataExtraida = text.split('\n')[4].split(' ')[2].replace('.', '/')
hora = text.split('\n')[10].split('-')[1]


print(hora)
