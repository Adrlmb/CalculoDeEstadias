transportadoras= ['FUTURO', 'G10', 'CARVALHO', 'FRIBON TRANSPORTES LTDA', 'D GRANEL', 'SIMOES BEBEDOURO', 'AGUETONI']
transportadoras.sort()

nome = 'fri'.upper()
nome = nome[0:3]



for i in range(len(transportadoras)):
    if nome in transportadoras[i]:
        print(transportadoras[i])






