
arq = open('Dados/Dados/DOMPNS2013.txt', 'r');

for linha in arq:
    print(len(linha))

arq.close()