
import random

long = 56

lista_test = [random.randint(1, 9)
              for i in range(long)]

print(lista_test)

rep = long // 10
# rest = long % 10

print(rep)

lista_listas = []

for x in range(rep):
    lista_cortada = lista_test[x*10: (10*(x+1))]
    lista_listas.append(lista_cortada)

lista_cortada = lista_test[rep*10:]
lista_listas.append(lista_cortada)

print(lista_listas)