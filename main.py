from openpyxl import load_workbook

# from openpyxl import Workbook
#
# # Cria uma nova planilha em branco
# wb = Workbook()
#
# # Seleciona a planilha ativa
# planilha = wb.active
#
# # Define o valor da célula A1
# planilha['A1'] = 'Olá, Excel com Python!'
#
# # Salva o arquivo
# wb.save('Exemplo.xlsx')


# from openpyxl import load_workbook
#
# # Carrega o arquivo existente
# wb = load_workbook('Exemplo.xlsx')
#
# # Seleciona a planilha ativa
# planilha = wb.active
#
# # Lê o valor da célula A1
# valor_celula = planilha['A1'].value
#
# # Imprime o valor na tela
# print(valor_celula)


# from openpyxl import load_workbook
#
# # Carrega o arquivo existente
# wb = load_workbook('Exemplo.xlsx')
#
# # Seleciona a planilha ativa
# planilha = wb.active
#
# # Atualiza o valor da célula A1
# planilha['A1'] = 'Novo valor'
#
# # Salva o arquivo
# wb.save('Exemplo.xlsx')


# Carrega o arquivo existente
wb = load_workbook('estadia\Cálculo estadia.xlsx')

# Seleciona a planilha ativa
planilha = wb.active

planilha['C3'] = 'Angico'  # Fornecedor
planilha['C4'] = '11844'  # NF
planilha['C5'] = 'Rocha UMA'  # Produto
planilha['B9'] = '11/09/2024 08:02'  # Data e Hora de chegada no formato DD/MM/AA HH:MM
planilha['B15'] = '20/09/2024 08:02'  # Data e Hora da saída no formato DD/MM/AA HH:MM
planilha['F3'] = 'G10'  # Transportador
planilha['F4'] = '654'  # Ct-e
planilha['F5'] = 'Jonivaldo Zequinha'  # Motorista
planilha['F12'] = '47,000'  # Peso NF com vírgula
planilha['E16'] = 'Motivo: Tombador quebrado'  # Motivo, vai se iniciar com "motivo:" e concatenar com o real motivo da estadia

# Lê o valor da célula A1
valor_celula = planilha['F13'].value

# Imprime o valor na tela
print(valor_celula)

wb.save('EstadiasCalculadas\Estadia_' + planilha[
    'F5'].value + '.xlsx')  # Salva a planilha com o nome Estadia + o nome do motorista


# Converte xlsx em pdf
# import  aspose.cells
#   from aspose.cells import Workbook
#   workbook = Workbook("input.xlsx")
#   workbook.save("Output.pdf")
