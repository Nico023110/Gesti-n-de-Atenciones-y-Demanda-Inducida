# -*- coding: utf-8 -*-
import os

# ======================================================
# RUTAS DEL PROYECTO
# ======================================================

BASE_DIR = r"C:\Proyecto_Poblacion"

RUTA_CATALOGOS = os.path.join(BASE_DIR, "datos", "catalogos")
RUTA_LOGS = os.path.join(BASE_DIR, "logs")

# Rutas dinámicas (se inicializan vacías y se llenan en main.py)
RUTA_NOMINALES = ""
RUTA_FEV = ""
RUTA_POBLACION = ""
RUTA_SALIDA = ""

def set_rutas_dinamicas(ruta_mes, eps_nombre):
    global RUTA_NOMINALES, RUTA_FEV, RUTA_POBLACION, RUTA_SALIDA
    
    RUTA_FEV = os.path.join(ruta_mes, eps_nombre, "fev")
    RUTA_NOMINALES = os.path.join(ruta_mes, eps_nombre, "nominal")
    RUTA_POBLACION = os.path.join(ruta_mes, eps_nombre, "poblacion")
    RUTA_SALIDA = os.path.join(ruta_mes, eps_nombre, "salida")

    # Crear carpetas si no existen
    os.makedirs(RUTA_NOMINALES, exist_ok=True)
    os.makedirs(RUTA_FEV, exist_ok=True)
    os.makedirs(RUTA_POBLACION, exist_ok=True)
    os.makedirs(RUTA_SALIDA, exist_ok=True)