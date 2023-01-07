from openpyxl import load_workbook
import logging
from os import name as osname

logging.basicConfig(filename="nomina.log",
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filemode='w')

logger = logging.getLogger()

# verificar tipo de sistema operativo
if osname == 'posix':   # linux, macOS
    wb = load_workbook('./input/Checkins and Checkouts (1).xlsx')
elif osname == 'nt':    # windows
    wb = load_workbook('.\\input\\Checkins and Checkouts (1).xlsx')
else:
    logger.critical(f'{osname} no es un valor valido de SO')


def hr2min(hora):   # convierte horas a minutos
    lista = [int(hora[:hora.find(':')]), int(hora[hora.find(':') + 1:-3])]
    if hora.find('AM') == -1:
        lista[0] = lista[0] + 12
    mins = lista[0] * 60 + lista[1]
    return mins


empleados = dict()
ids = list()
sheet = wb['Invoice']

# obtener dimensiones de hoja
num_empleados = sheet.calculate_dimension()
# cantidad de instancias (de checkin o checkout)
num_empleados = int(num_empleados[(2+num_empleados.find(':')):]) - 5
# loggear en caso de falla


class Empleado:
    """Empleado con número, ID, nombre, instancias de entrada o salida,
    cantidad de retardos y salidas anticipadas"""

    def __init__(self, num, id_nom, nombre):
        self.num, self.id_nom, self.nombre = num, id_nom, nombre
        self.instancias = list()    # instancias de entrada o salida
        self.retardos = list()      # retardos
        self.anticipadas = list()   # salidas anticipadas
        self.trabajado = list()     # tiempo trabajado

    def __str__(self):
        x = (f'Empleado:\n   num: {self.num}\n   id: {self.id_nom}\n'
             f'   nombre: {self.nombre}\n'
             f'   retardos: {self.retardos}\n'
             f'   salidas anticipadas: {self.anticipadas}\n'
             f'   trabajado: {self.trabajado}\n'
             f'   instancias:\n')
        for i in range(len(self.instancias)):
            x += '    ' + str(self.instancias[i]) + '\n'
        return x


# crear lista de empleados
for empleado in range(6, num_empleados + 5):
    # obtener valores
    emp_num = int(sheet['A'+str(empleado)].value)    # número de empleado
    emp_id = sheet['B'+str(empleado)].value     # id de nómina de empleado
    emp_nom = sheet['C'+str(empleado)].value    # nombre de empleado

    # agregar objeto de empleado si no está
    if sheet['A'+str(empleado)].value not in ids:
        empleados.update({emp_num: Empleado(emp_num, emp_id, emp_nom)})
        ids.append(sheet['A'+str(empleado)].value)

# añadir checkins y checkouts
    for i in empleados:
        if int(emp_num) == i:
            empleados[int(emp_num)].instancias.append({
                'entrada':  sheet['H'+str(empleado)].value,
                'salida':   sheet['I'+str(empleado)].value,
                'checkin':  sheet['J'+str(empleado)].value,
                'checkout': sheet['K'+str(empleado)].value,
                'fecha':    sheet['F'+str(empleado)].value})


# obtener horas trabajadas, inasistencias, retardos...
for empleado in empleados:
    for instancia in empleados[empleado].instancias:
        entr = hr2min(instancia['entrada'])
        sali = hr2min(instancia['salida'])
        try:
            chin = hr2min(instancia['checkin'])
        except AttributeError:
            chin = None
            # Si no hay checkin para la instancia
            logger.warning(f'Sin checkin ({instancia["checkin"]})')
        try:
            chou = hr2min(instancia['checkout'])
        except AttributeError:
            # Si no hay checkout para la instancia
            chou = None
            logger.warning(f'Sin checkout ({instancia["checkout"]})')

        if chou is None or chin is None:  # tiempo trabajado
            logger.error(f'AttributeError: Error calculando tiempo trabajado'
                         f'\n   Checkin:     {instancia["checkin"]} = {chin}'
                         f'\n   Checkout:    {instancia["checkout"]} = {chou}')

        else:
            empleados[empleado].trabajado.append(chou - chin)
            # agregar retardo
            if entr > sali:
                empleados[empleado].retardos.append(chin)

            # agregar salida anticipada
            if sali < chou:
                empleados[empleado].retardos.append(chou)

# Escribir datos
wb.create_sheet('Reporte')
for empleado in empleados:
    # añadir a la hoja general

    # crear una hoja por empleado
    titulo = empleados[empleado].nombre.find(' ')
    titulo = empleados[empleado].nombre[:titulo]
    wb.create_sheet(title=titulo)
    sheet = wb[titulo]
    insts = empleados[empleado].instancias

    sheet['A1'], sheet['B1'] = 'Nombre', empleados[empleado].nombre
    sheet['C1'], sheet['D1'] = 'ID', empleados[empleado].id_nom
    sheet['E1'], sheet['F1'] = 'Num', empleados[empleado].num

    sheet['A2'] = 'Hora de entrada'
    sheet['B2'] = 'Checkin'
    sheet['C2'] = 'Hora de salida'
    sheet['D2'] = 'Checkout'
    sheet['E2'] = 'Fecha'

    for i, instancia in enumerate(insts):
        sheet['A' + str(i+3)] = instancia['entrada']
        sheet['B' + str(i+3)] = instancia['checkin']
        sheet['C' + str(i+3)] = instancia['salida']
        sheet['D' + str(i+3)] = instancia['checkout']
        sheet['E' + str(i+3)] = instancia['fecha']

if osname == 'posix':   # linux, macOS
    wb.save('./output/Checkins and Checkouts.xlsx')
elif osname == 'nt':    # windows
    wb.save('.\\output\\Checkins and Checkouts.xlsx')
