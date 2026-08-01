# -*- coding: utf-8 -*-
"""
Módulo: cargar_nominal.py
Descripción: Carga todas las nominales/cohortes desde datos/nominales,
             normaliza sus columnas y les asigna el nombre de la EAPB.
"""

import os
import pandas as pd
from utilidades import separador
from normalizar_columnas import normalizar_dataframe
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config')))
import config


def extraer_nombre_eapb(nombre_archivo: str) -> str:
    """
    Extrae un nombre limpio de la EAPB a partir del nombre del archivo.
    """
    nombre_upper = nombre_archivo.upper()
    
    if 'EMSSANAR' in nombre_upper:
        return 'EMSSANAR'
    elif 'NUEVA' in nombre_upper or 'EPS' in nombre_upper:
        return 'NUEVA EPS'
    elif 'SANITAS' in nombre_upper:
        return 'SANITAS'
    elif 'ASMET' in nombre_upper:
        return 'ASMET SALUD'
    else:
        return os.path.splitext(nombre_archivo)[0]


def cargar_nominales() -> pd.DataFrame:
    """
    Lee todas las cohortes/nominales de la carpeta datos/nominales/,
    les inyecta la columna 'eapb' y las consolida en un único DataFrame.
    """
    separador("LEYENDO NOMINALES")

    if not os.path.exists(config.RUTA_NOMINALES):
        print(f" [ERROR CRÍTICO] No se encontró la carpeta de nominales en: {config.RUTA_NOMINALES}")
        return None

    archivos = [f for f in os.listdir(config.RUTA_NOMINALES) if f.endswith(('.xlsx', '.xls', '.csv'))]

    if not archivos:
        print(f" [ADVERTENCIA] No se encontraron archivos en: {config.RUTA_NOMINALES}")
        return None

    lista_nominales = []

    for archivo in archivos:
        ruta_completa = os.path.join(config.RUTA_NOMINALES, archivo)
        print(f"Leyendo: {archivo}")

        if archivo.endswith('.csv'):
            df = pd.read_csv(ruta_completa, low_memory=False)
        else:
            df = pd.read_excel(ruta_completa)

        # Normalizar encabezados y traducir alias
        df, _ = normalizar_dataframe(df)

        # Asignar únicamente la columna 'eapb' (se removió 'archivo_origen')
        df = df.copy()
        nombre_eapb = extraer_nombre_eapb(archivo)
        df['eapb'] = nombre_eapb

        print(f"   Pacientes: {len(df):,} | EAPB identificada: {nombre_eapb}")
        lista_nominales.append(df)

    # Consolidar todas las nominales
    df_nominal = pd.concat(lista_nominales, ignore_index=True)

    print("-" * 80)
    print(f"Total pacientes consolidados : {len(df_nominal):,}")
    print(f"Total columnas final         : {len(df_nominal.columns)}")
    print("-" * 80)

    return df_nominal