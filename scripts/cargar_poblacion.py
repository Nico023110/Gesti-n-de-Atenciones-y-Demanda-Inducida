# -*- coding: utf-8 -*-
"""
Módulo: cargar_poblacion.py
Descripción: Carga la base de datos de población asignada a la IPS desde datos/poblacion/,
             normaliza columnas e identifica la EAPB de cada archivo.
"""

import os
import pandas as pd
from utilidades import separador
from normalizar_columnas import normalizar_dataframe
from cargar_nominal import extraer_nombre_eapb
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config')))
import config


import numpy as np


def _preparar_demografia_poblacion(df_poblacion: pd.DataFrame) -> pd.DataFrame:
    """
    Enriquece la base de población asignada a la IPS con nombre_completo,
    edad_meses y unificación de curso_vida para permitir su evaluación en el motor de reglas.
    """
    df = df_poblacion.copy()

    # Construir nombre completo
    p_nom = df.get('primer_nombre', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    s_nom = df.get('segundo_nombre', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    p_ape = df.get('primer_apellido', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    s_ape = df.get('segundo_apellido', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    nombre_completo = (p_nom + ' ' + s_nom).str.strip() + ' ' + (p_ape + ' ' + s_ape).str.strip()
    df['nombre_completo'] = nombre_completo.str.strip().str.upper()

    # Unificar curso_vida / ciclovida
    if 'curso_vida' not in df.columns and 'ciclovida' in df.columns:
        df['curso_vida'] = df['ciclovida']
    elif 'curso_vida' in df.columns and 'ciclovida' in df.columns:
        df['curso_vida'] = df['curso_vida'].fillna(df['ciclovida'])

    # Calcular edad en meses y días a partir de fecha_nacimiento
    if 'fecha_nacimiento' in df.columns:
        hoy = pd.Timestamp.now()
        nacimiento = pd.to_datetime(df['fecha_nacimiento'], errors='coerce', dayfirst=True)

        edad_meses = (hoy.year - nacimiento.dt.year) * 12 + (hoy.month - nacimiento.dt.month)
        dias_insuficientes = hoy.day < nacimiento.dt.day
        edad_meses = np.where(dias_insuficientes, edad_meses - 1, edad_meses)

        edad_dias = (hoy - nacimiento).dt.days

        df['edad_meses'] = pd.Series(edad_meses, index=df.index).fillna(-1).astype(int)
        df['edad_dias'] = pd.Series(edad_dias, index=df.index).fillna(-1).astype(int)
    else:
        df['edad_meses'] = -1
        df['edad_dias'] = -1

    return df


def cargar_poblacion() -> pd.DataFrame:
    """
    Lee todos los archivos de población asignada desde la carpeta datos/poblacion/,
    normaliza sus columnas, calcula la edad en meses y enriquece la demografía.
    Retorna un DataFrame consolidado con todos los pacientes asignados a la IPS.
    """
    separador("LEYENDO POBLACIÓN ASIGNADA A LA IPS")

    if not os.path.exists(config.RUTA_POBLACION):
        print(f" [ERROR CRÍTICO] No se encontró la carpeta de población en: {config.RUTA_POBLACION}")
        return None

    archivos = [f for f in os.listdir(config.RUTA_POBLACION) if f.endswith(('.xlsx', '.xls', '.csv'))]

    if not archivos:
        print(f" [ADVERTENCIA] No se encontraron archivos en: {config.RUTA_POBLACION}")
        return None

    lista_poblacion = []

    for archivo in archivos:
        ruta_completa = os.path.join(config.RUTA_POBLACION, archivo)
        print(f"Leyendo: {archivo}")

        try:
            if archivo.endswith('.csv'):
                df = pd.read_csv(ruta_completa, low_memory=False)
            else:
                # Leer TODAS las hojas del Excel (sheet_name=None retorna un dict)
                dict_hojas = pd.read_excel(ruta_completa, sheet_name=None)
                hojas_parciales = []
                for nombre_hoja, df_hoja in dict_hojas.items():
                    if 'AFILIADO' in nombre_hoja.upper():
                        print(f"   Hoja seleccionada: {nombre_hoja} -> {len(df_hoja):,} registros")
                        hojas_parciales.append(df_hoja)
                    else:
                        print(f"   Hoja ignorada: {nombre_hoja}")
                
                if hojas_parciales:
                    df = pd.concat(hojas_parciales, ignore_index=True)
                else:
                    print(f" [ADVERTENCIA] No se encontro hoja de AFILIADOS en {archivo}")
                    continue
        except Exception as e:
            print(f" [ERROR] No se pudo leer {archivo}: {e}")
            continue

        # Normalizar encabezados y traducir alias
        df, _ = normalizar_dataframe(df)

        # Asignar EAPB desde el nombre del archivo
        df = df.copy()
        nombre_eapb = extraer_nombre_eapb(archivo)
        df['eapb'] = nombre_eapb

        print(f"   Registros: {len(df):,} | EAPB identificada: {nombre_eapb}")
        lista_poblacion.append(df)

    if not lista_poblacion:
        print(" [ADVERTENCIA] No se pudo cargar ningún archivo de población.")
        return None

    # Consolidar toda la población y eliminar duplicados por documento (conservando el más reciente)
    df_poblacion = pd.concat(lista_poblacion, ignore_index=True)
    if 'nro_identificacion' in df_poblacion.columns:
        filas_antes = len(df_poblacion)
        df_poblacion['nro_identificacion'] = df_poblacion['nro_identificacion'].astype(str).str.strip().str.upper()
        df_poblacion = df_poblacion.drop_duplicates(subset=['nro_identificacion'], keep='last')
        print(f" -> Deduplicación de población: de {filas_antes:,} a {len(df_poblacion):,} usuarios únicos.")

    # Preparar demografía (edad_meses, nombre_completo, curso_vida)
    df_poblacion = _preparar_demografia_poblacion(df_poblacion)

    print("-" * 80)
    print(f"Total registros población : {len(df_poblacion):,}")
    print(f"Total columnas            : {len(df_poblacion.columns)}")
    print("-" * 80)

    return df_poblacion

