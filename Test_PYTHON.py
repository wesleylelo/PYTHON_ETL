cpf = '075.694.495-33'
lista1 = cpf.split('.')
valor = lista1[2].split('-')
lista1[2] = valor[0]
print(lista1)
string = lista1[0] + lista1[1] + lista1[2]
print(string)
soma = 0
for i in range(10, 1, -1):
    soma += int(i)*int(string[10-i])
print(soma)
g = 0
if 11 - soma%11 > 3:
    g = 3
else: 
    g = 11-soma%11
string += str(g)
print(string)