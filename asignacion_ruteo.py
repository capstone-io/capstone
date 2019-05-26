import xlrd

from gurobipy import *
import networkx as nx #Para trabajar con grafos
import matplotlib.pyplot as plt #Para visualizar los grafos

from openpyxl import Workbook, load_workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font,Alignment
from datetime import date, timedelta
import os

from asignacion_beto import *

def repartir_turnos(planta, dia, obras_asignadas):
    repartición_turnos = {}
    turno1 = 0
    turno2 = 0
    turno3 = 0
    for turno in turnos:
        repartición_turnos[turno] = []
    for obra in obras_asignadas:
        if turnos_ab[obra, 3] == 0 and turnos_ab[obra, 2] == 0:
            repartición_turnos[1].append(obra)
            turno1 += demanda_diaria[obra, dia]
        elif turnos_ab[obra, 1] == 0 and turnos_ab[obra, 2] == 0:
            repartición_turnos[3].append(obra)
            turno3 += demanda_diaria[obra, dia]
        elif turnos_ab[obra, 3] == 0 and turnos_ab[obra, 1] == 0:
            repartición_turnos[2].append(obra)
            turno2 += demanda_diaria[obra, dia]
        elif turnos_ab[obra, 3] == 0:
            if turno1 >= turno2:
                repartición_turnos[2].append(obra)
                turno2 += demanda_diaria[obra, dia]
            else:
                repartición_turnos[1].append(obra)
                turno1 += demanda_diaria[obra, dia]
        elif turnos_ab[obra, 2] == 0:
            if turno1 >= turno3:
                repartición_turnos[3].append(obra)
                turno3 += demanda_diaria[obra, dia]
            else:
                repartición_turnos[1].append(obra)
                turno1 += demanda_diaria[obra, dia]
        elif turnos_ab[obra, 1] == 0:
            if turno2 >= turno3:
                repartición_turnos[3].append(obra)
                turno3 += demanda_diaria[obra, dia]
            else:
                repartición_turnos[2].append(obra)
                turno2 += demanda_diaria[obra, dia]
        else:
            if turno1 <= turno2 and turno1<=turno3:
                repartición_turnos[1].append(obra)
                turno1 += demanda_diaria[obra, dia]
            elif turno2 <= turno1 and turno2<=turno3:
                repartición_turnos[2].append(obra)
                turno2 += demanda_diaria[obra, dia]
            elif turno3 <= turno1 and turno3<=turno2:
                repartición_turnos[3].append(obra)
                turno3 += demanda_diaria[obra, dia]
    return repartición_turnos
print(repartir_turnos(1, 1, obras_asig_1_1))

reparticion_turnos = {}

for planta in plantas:
    for dia in dias:
        asignacion_turnos = repartir_turnos(planta, dia, obras_asignadas[planta, dia])
        for turno in asignacion_turnos:
            reparticion_turnos[planta, dia, turno] = asignacion_turnos[turno]
            
print(reparticion_turnos)

demanda_prueba = {} 
for planta in plantas:
    for dia in dias:
        for turno in turnos:
            demanda_prueba[planta, dia, turno] = 0
            for obra in reparticion_turnos[planta, dia, turno]:
                demanda_prueba[planta, dia, turno] += demanda_diaria[obra, dia]

print('dem', demanda_prueba)
