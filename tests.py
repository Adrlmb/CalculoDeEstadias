import pdfplumber

pdf = pdfplumber.open('tickets/belocal - cal.pdf')
page = pdf.pages[0]
text = page.extract_text()

transportadoras = ['MINERACAO BELOCAL', 'CARVALHO TRANSPORTES', 'FRIBON TRANSPORTES', 'FUTURO LOGISTICA',
                   'SIMOES BEBEDOURO', 'TRANSLOPES TRANSPORTES']

codigoTransportadoras = ['2000227172', '2000224719', '2000215499', '2000226886', '204005', '207327']

referencia = text.split('\n')[10].split('-')[0].split(' ')[1]
print(referencia)

for i in range(len(transportadoras)):
    if codigoTransportadoras[i] == referencia:
        print(transportadoras[i])
