from openpyxl import load_workbook
import logging
from os import name as osname

logging.basicConfig(filename="nomina.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')

logger = logging.getLogger()

# verificar tipo de sistema operativo
if osname == 'posix':   # linux, macOS
    wb = load_workbook('./input/Checkins and Checkouts (1).xlsx')
elif osname == 'nt':    # windows
    wb = load_workbook('.\\input\\Checkins and Checkouts (1).xlsx')
else:
    logger.error(f'{osname} no es un valor valido de SO')   # log


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
        return (f'Empleado:\n   num: {self.num}\n   id: {self.id_nom}\n'
                f'   nombre: {self.nombre}\n'
                f'   instancias: {self.instancias}\n'
                f'   retardos: {self.retardos}\n'
                f'   salidas anticipadas: {self.anticipadas}')


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
                'checkout': sheet['K'+str(empleado)].value})


# obtener horas trabajadas, inasistencias, retardos...
for empleado in empleados:
    for instancia in empleados[empleado].instancias:
        try:
            entr = hr2min(instancia['entrada'])
            sali = hr2min(instancia['salida'])
            chin = hr2min(instancia['checkin'])
            chou = hr2min(instancia['checkout'])
        except AttributeError:
            pass    # Si no hay checkin o checkout para la instancia

        try:  # tiempo trabajado
            empleados[empleado].trabajado.append(chou - chin)
        except AttributeError:
            logger.error(f'AttributeError: Error calculando tiempo trabajado'
                         f'\n   Checkin:     {instancia["checkin"]} = {chin}'
                         f'\n   Checkout:    {instancia["checkout"]} = {chou}')

        try:
            # agregar retardo
            if entr > sali:
                empleados[empleado].retardos.append(chin)
        except AttributeError:
            logger.error(f' AttributeError: Error calculando retardo'
                         f'\n   Entrada:     {instancia["entrada"]}'
                         f'\n   Checkin:     {instancia["checkin"]}')

        try:
            # agregar salida anticipada
            if sali < chou:
                empleados[empleado].retardos.append(chou)

        except AttributeError:
            logger.error(f'AttributeError: Error calculando tiempo salida'
                         f'\n   Empleado:    {instancia["salida"]}'
                         f'\n   Instancia:   {instancia["checkout"]}')
