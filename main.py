import pdfplumber
from openpyxl import load_workbook

pdf = pdfplumber.open('estadia\TICKET 1.pdf')
page = pdf.pages[0]
text = page.extract_text()
# print(text.split('\n')[7].split(':')[1])

# Carrega o arquivo existente
wb = load_workbook('estadia\Cálculo estadia.xlsx')

# Seleciona a planilha ativa
planilha = wb.active


def dataDeSaida():
    data = text.split('\n')[4].split(' ')[2].replace('.', '/')
    hora = text.split('\n')[4].split(' ')[3].split(':')[0]
    minutos = text.split('\n')[4].split(' ')[3].split(':')[1]

    return data + ' ' + hora + ':' + minutos


planilha['C3'] = 'INPUT'  # Fornecedor
planilha['C4'] = int(text.split('\n')[7].split(':')[1])  # NF
planilha['C5'] = text.split('\n')[9].split('-')[1]  # Produto, impor condição se for Rocha!!
planilha['B9'] = 'INPUT'  # Data e Hora de chegada no formato DD/MM/AA HH:MM
planilha['B15'] = dataDeSaida()  # Data e Hora da saída no formato DD/MM/AA HH:MM
planilha['F3'] = 'G10'  # Transportador
planilha['F4'] = 'INPUT'  # Ct-e
planilha['F5'] = 'INPUT'  # Motorista
planilha['F12'] = text.split('\n')[5].split(": ")[1]  # Peso NF com vírgula
planilha['E16'] = 'INPUT'  # Motivo, vai se iniciar com "motivo:" e concatenar com o real motivo da estadia

# Lê o valor da célula A1
valor_celula = planilha['C4'].value

# Imprime o valor na tela
print(valor_celula)

wb.save('EstadiasCalculadas\Estadia_0.002.xlsx')  # Salva a planilha com o nome Estadia + o nome do motorista

#####################################################################################################


# Converte xlsx em pdf
# import  aspose.cells
#   from aspose.cells import Workbook
#   workbook = Workbook("input.xlsx")
#   workbook.save("Output.pdf")
