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
from funciones_auxiliares import distancia_minima, condiciones_camion, tiempo_puntos, volver_planta

sheet_camiones = wb.sheet_by_index(2)

# camiones dict = [posicion_actual, horas_actual, capacidad_actual, hora_actual, trayectos[dia, turno] = (planta,obra)]
camiones_chicos = int(sheet_camiones.cell_value(7, 3))
camiones_medianos = int(sheet_camiones.cell_value(8, 3))
camiones_grandes = int(sheet_camiones.cell_value(9, 3))

camion_capacidad = {}


camiones = {}

for camion_chico in range(camiones_chicos):
    camiones[camion_chico+1] = [(0,0), 0, int(sheet_camiones.cell_value(7, 2)), 
                   int(sheet_camiones.cell_value(7, 4)), {}]
    camion_capacidad[camion_chico+1] = int(sheet_camiones.cell_value(7, 2))

for camion_mediano in range(camiones_chicos, camiones_chicos + camiones_medianos):
    camiones[camion_mediano+1] = [(0,0), 0, int(sheet_camiones.cell_value(8, 2)), 
                   int(sheet_camiones.cell_value(8, 4)), {}]
    
    camion_capacidad[camion_mediano+1] = int(sheet_camiones.cell_value(8, 2))

    
for camion_grande in range(camiones_chicos + camiones_medianos, camiones_chicos + camiones_medianos + camiones_grandes):
    camiones[camion_grande+1] = [(0,0), 0, int(sheet_camiones.cell_value(9, 2)), 
                   int(sheet_camiones.cell_value(9, 4)), {}]
    camion_capacidad[camion_grande+1] = int(sheet_camiones.cell_value(9, 2))


for camion in camiones:
    for dia in dias:
        for turno in turnos:
            camiones[camion][4][dia, turno] = []




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


reparticion_turnos = {}

for planta in plantas:
    for dia in dias:
        asignacion_turnos = repartir_turnos(planta, dia, obras_asignadas[planta, dia])
        for turno in asignacion_turnos:
            reparticion_turnos[planta, dia, turno] = asignacion_turnos[turno]
            

demanda_prueba = {} 
for planta in plantas:
    for dia in dias:
        for turno in turnos:
            demanda_prueba[planta, dia, turno] = 0
            for obra in reparticion_turnos[planta, dia, turno]:
                demanda_prueba[planta, dia, turno] += demanda_diaria[obra, dia]
                
# dict con las demandas incumplidas
demandas_incumplidas = demanda_diaria.copy()
inventario_restante = {}
for planta in plantas:
    for dia in dias:
        inventario_restante[planta, dia] = inventario_plantas[planta, dia]

print(inventario_restante)
        
for dia in reversed(dias):
    for turno in turnos:
        for planta in plantas:
            lista_obras = reparticion_turnos[planta, dia, turno]
            for camion in camiones:
                # ya no se puede satisfacer ninguna obra
                if inventario_restante[planta, dia] == 0:
                    #print('entra')
                    continue
                #print('camion numero', camion)
                # lista de obras a las que puede ir el camion
                lista_obras_camion = lista_obras.copy()
                camiones[camion][0] = posiciones_plantas[planta]
                # se resta del inventario de la planta lo que saca el camion
                if inventario_restante[planta, dia] > camiones[camion][2]:
                    inventario_restante[planta, dia] -= camiones[camion][2]
                else:
                    camiones[camion][2] = inventario_restante[planta, dia]
                    inventario_restante[planta, dia] = 0
                # condicion: revisa si tiene tiempo para ir a otra obra 
                # o capacidad > 0 o tiene que volver a planta por tiempo
                while True:
                    # revisa que el camion este cumpliendo las condiciones para
                    # seguir funcionando en ese turno (tiempo, capacidad)
                    if condiciones_camion(camiones[camion], lista_obras_camion) is False:
                        break
                    # diccionario con las distacias de las obras asignadas a esa 
                    # planta en ese dia en ese turno
                    dict_distancias = { obra: posiciones_obras[obra] for obra in lista_obras_camion }
                    obra_destino = distancia_minima(camiones[camion][0], dict_distancias)
                    #revisa si al ir a esa obra alcanza a volver a la planta
                    if volver_planta(dict_distancias[obra_destino], posiciones_plantas[planta],
                                     camiones[camion][1]) is False:
                        #print('false')
                        lista_obras_camion.remove(obra_destino)
                        continue
                    # se cambia de obra
                    else:
                        #print('true')
                        # agrega trayecto a la lista de trayectos del camion
                        camiones[camion][4][dia, turno].append((camiones[camion][0], dict_distancias[obra_destino]))
                        # le suma al camion el tiempo de recorrido
                        camiones[camion][1] += tiempo_puntos(camiones[camion][0], dict_distancias[obra_destino])
                        # actualiza posicion actual a la de la obra de destino
                        camiones[camion][0] = dict_distancias[obra_destino]
                        # si la demanda que queda por cumplir de esa obra es mayor a la
                        # capacidad del camion se le resta a lo que falta
                        # la capacidad actual del camion queda en 0 y debe volver a la planta
                        if demandas_incumplidas[obra_destino, dia] > camiones[camion][2]:
                            demandas_incumplidas[obra_destino, dia] -= camiones[camion][2]
                            camiones[camion][2] = 0
                            # agrega el trayecto a la lista de trayectos de ese dia en ese turno
                            camiones[camion][4][dia, turno].append((dict_distancias[obra_destino], posiciones_plantas[planta]))
                            # le suma el tiempo recorrido al camion
                            camiones[camion][1] += tiempo_puntos(camiones[camion][0], posiciones_plantas[planta])
                            # vuelve a la planta
                            camiones[camion][0] = posiciones_plantas[planta]
                            # se rellena el camion y se le saca inventario a la planta
                            if inventario_restante[planta, dia] > camion_capacidad[camion]:
                                #print('entra 3', inventario_restante[planta, dia], camion_capacidad[camion])
                                inventario_restante[planta, dia] -= camiones[camion][2]
                                camiones[camion][2] = camion_capacidad[camion]
                            else:
                                #
                                print('entra 2')
                                camiones[camion][2] = inventario_restante[planta, dia]
                                inventario_restante[planta, dia] = 0 
                                break
                            #se elimina la obra de la lista de obras para ese camion
                            lista_obras_camion.remove(obra_destino)
                            continue
                        # la capacidad del camion es mayor a la demanda de esa obra
                        else:
                            # se le resta a la capacidad del camion la demanda incumplida de esa obra
                            camiones[camion][2] -= demandas_incumplidas[obra_destino, dia]
                            # se borra la obra de la lista de obras que puede cumplir el camion
                            # y de la lista de obras asignadas
                            lista_obras_camion.remove(obra_destino)
                            lista_obras.remove(obra_destino)
                            continue
                # se reincian parametros del camion para el turno siguiente
                #tiempo recorrido=0 , capacidad = capacidad maxima
                camiones[camion][1] = 0
                camiones[camion][2] = camion_capacidad[camion]
              
for 

#print('inventario asignaciones', inventario_plantas)                      
#print('demandas', demandas_incumplidas)
#print('inventario_restante', inventario_restante)
#print(camiones)
                        
                    
            #print(dia, planta, turno, lista_obras)

