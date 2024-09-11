from openpyxl import load_workbook

# Carrega o arquivo existente
wb = load_workbook('estadia\Cálculo estadia.xlsx')

# Seleciona a planilha ativa
planilha = wb.active

# Lê o valor da célula A1
valor_celula = planilha['B3'].value

# Imprime o valor na tela
print(valor_celula)

planilha['B3'] = 'XERECA'
valor_celula = planilha['B3'].value

print(valor_celula)

wb.save('Exemplo1.xlsx')

