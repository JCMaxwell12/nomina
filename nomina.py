from openpyxl import load_workbook
import logging

wb = load_workbook('./input/Checkins and Checkouts (1).xlsx')


def hr2min(hora):   # convierte horas a minutos
    lista = [int(hora[:hora.find(':')]), int(hora[hora.find(':') + 1:-3])]
    if hora.find('AM') == -1:
        lista[0] = lista[0] + 12
    mins = lista[0] * 60 + lista[1]
    return mins

logging.basicConfig(filename="nomina.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
 
logger = logging.getLogger()


empleados = dict()
ids = list()
sheet = wb['Invoice']

num_empleados = sheet.calculate_dimension()
num_empleados = int(num_empleados[(2+num_empleados.find(':')):]) - 5
# loggear en caso de falla


class Empleado:
    """Empleado con número, ID y nombre"""

    def __init__(self, num, id_nom, nombre):
        self.num, self.id_nom, self.nombre = num, id_nom, nombre
        self.instancias = list()     # instancias de entrada o salida
        self.retardos = list()
        self.anticipadas = list()   # salidas anticipadas

    def __str__(self):
        return f'Empleado({self.num}, {self.id_nom}, {self.nombre}, {self.instancias})'


# crear lista de empleados
for empleado in range(6, num_empleados + 5):
    # obtener valores
    emp_num = sheet['A'+str(empleado)].value    # número de empleado
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
    for instancia in empleado.instanicas:
        try:  # tiempo trabajado
            hr2min(instancia['checkin']) - hr2min(instancia['checkout'])

            # agregar retardo
            if hr2min(instancia['entrada']) > hr2min(instancia['checkin']):
                pass
        except TypeError:
            empleado.retardos.append(instancia['checkin'])

        try:
            # agregar salida anticipada
            if hr2min(instancia['salida']) < hr2min(instancia['checkout']):
                pass

        except TypeError:
            empleado.retardos.append(instancia['checkin'])
