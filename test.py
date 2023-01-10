#!/sbin/python
import os
try:
    os.remove('output/*.xlsx')
except:
    print('output/Checkins and Checkouts.xlsx no existe')
import nomina

print(f'dict empleados: {nomina.empleados}\n\nids: {nomina.ids}\n')

for i in nomina.empleados:
    print(nomina.empleados[i], '\n')

print(' ##### Log #####')
with open('nomina.log', 'r') as log:
    print(log.read())
