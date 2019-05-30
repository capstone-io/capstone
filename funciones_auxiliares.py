from math import*

def distancia_manhattan(x,y):
    #print('x', x)
    #print('y', y)

    return sum(abs(a-b) for a,b in zip(x,y))

def distancia_minima(punto, distancias):   
    # returno obra mas cercana a un punto dado    
    return min(distancias.items(), key=lambda x: distancia_manhattan(punto, x[1]))[0]

def condiciones_camion(camion, obras):
    # ya no hay obras que pueda suministrar el camion
    if len(obras) == 0:
        return False
    # camion con capacidad que no alcanza para mas
    elif camion[2] <= 0:
        #print('capcidad', camion[1])
        return False
    
def tiempo_puntos(x, y):
    return distancia_manhattan(x,y) / 400 # tiempo con 40 km/h

    
def volver_planta(punto, coord_planta_actual, tiempo_camion):
    # condicion si el tiempo actual que lleva el camion en ese turno mas el tiempo
    # que se demoraria en llegar a la planta es mayor al tiempo del turno
    # para que vuelva a la planta al final del turno
    #print('tiempos', tiempo_puntos(punto, coord_planta_actual), tiempo_camion)
    if tiempo_puntos(punto, coord_planta_actual) + tiempo_camion > 8:
        return False
    else:
        return True
    
