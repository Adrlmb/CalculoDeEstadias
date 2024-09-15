nome = ''
sobrenome = ''
ultimoNome = ''

campos = [nome, sobrenome, ultimoNome]

for i in range(3):
    if campos[i] in '':
        campos[i] = input('Digite um nome: ')

    print(campos[i])


print(campos)