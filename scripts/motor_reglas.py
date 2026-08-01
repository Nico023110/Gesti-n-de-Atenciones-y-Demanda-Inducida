# -*- coding: utf-8 -*-
"""
Created on Jul 2026

@author: analisisdedatos
Módulo: motor_reglas.py
Descripción: Evalúa las reglas de negocio en dos fases:
             1. Determina qué actividades necesita cada paciente (por edad, sexo).
             2. Verifica en la FEV cuáles de esas actividades ya fueron realizadas.
             Retorna las actividades realizadas y las pendientes.
"""

import os
import pandas as pd
# pyrefly: ignore [missing-import]
import numpy as np
from utilidades import separador
from config import RUTA_CATALOGOS

# Ruta absoluta segura utilizando config.py (evita errores de barras invertidas en Windows)
RUTA_REGLAS = os.path.join(RUTA_CATALOGOS, "reglas_actividades.xlsx")
HOJA_REGLAS = "REGLAS"


def cargar_reglas(ruta_excel: str = RUTA_REGLAS, hoja: str = HOJA_REGLAS) -> pd.DataFrame:
    """
    Lee el Excel de reglas, limpia encabezados y filtra solo las reglas activas.
    Se lee como texto (dtype=str) para evitar que códigos CUPS se lean como flotantes.
    """
    if not os.path.exists(ruta_excel):
        print(f" [ERROR CRÍTICO] No se encontró el catálogo de reglas en: {ruta_excel}")
        return pd.DataFrame()

    df_reglas = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)
    
    # Normalizar encabezados a mayúsculas y sin espacios
    df_reglas.columns = df_reglas.columns.str.strip().str.upper()

    # Filtrar únicamente las reglas activas (SI / 1 / TRUE)
    if 'ACTIVO' in df_reglas.columns:
        df_reglas = df_reglas[df_reglas['ACTIVO'].astype(str).str.strip().str.upper().isin(['SI', '1', 'TRUE'])].copy()

    print(f" -> Catálogo cargado: {len(df_reglas)} reglas activas.")
    return df_reglas


def _evaluar_condicion(serie_fev: pd.Series, valor_regla) -> pd.Series:
    """
    Evalúa coincidencias exactas o por listas separadas por comas en una celda del Excel.
    Ejemplo: Si en el Excel escribes "890201, 890202", evalúa si la FEV tiene cualquiera de los dos.
    """
    if pd.isna(valor_regla) or str(valor_regla).strip() == '' or str(valor_regla).strip().upper() in ['NAN', 'NONE', 'NAT']:
        return pd.Series(True, index=serie_fev.index)

    lista_valores = [v.strip().upper() for v in str(valor_regla).split(',') if v.strip()]
    lista_valores = [v[:-2] if v.endswith('.0') else v for v in lista_valores]

    serie_clean = serie_fev.astype(str).str.strip().str.upper()
    serie_clean = serie_clean.str.replace(r'\.0$', '', regex=True)

    return serie_clean.isin(lista_valores)


def _filtrar_fev_por_regla(df_fev_sub: pd.DataFrame, regla: pd.Series) -> pd.DataFrame:
    """
    Filtra registros FEV que cumplen los criterios CLÍNICOS de una regla
    (CUPS, Finalidad, Diagnósticos). NO aplica filtros demográficos aquí.
    """
    mask = pd.Series(True, index=df_fev_sub.index)

    # Filtro: CUPS (busca en codprocedimiento O codconsulta)
    if 'CUPS' in regla and not pd.isna(regla['CUPS']) and str(regla['CUPS']).strip() != '':
        mask_cups = pd.Series(False, index=df_fev_sub.index)
        if 'codprocedimiento' in df_fev_sub.columns:
            mask_cups |= _evaluar_condicion(df_fev_sub['codprocedimiento'], regla['CUPS'])
        if 'codconsulta' in df_fev_sub.columns:
            mask_cups |= _evaluar_condicion(df_fev_sub['codconsulta'], regla['CUPS'])
        mask &= mask_cups

    # Filtro: Finalidad
    if 'FINALIDAD' in regla and not pd.isna(regla['FINALIDAD']) and str(regla['FINALIDAD']).strip() != '':
        if 'finalidadtecnologiasalud' in df_fev_sub.columns:
            mask &= _evaluar_condicion(df_fev_sub['finalidadtecnologiasalud'], regla['FINALIDAD'])

    # Filtro: Diagnósticos (busca en cualquiera de las 4 columnas de DX del RIPS)
    cols_dx_regla = ['DX_PRINCIPAL', 'DX_RELACIONADO', 'DX_RELACIONADO1', 'DX_RELACIONADO2']
    hay_filtro_dx = any(col in regla and not pd.isna(regla[col]) and str(regla[col]).strip() != '' for col in cols_dx_regla)

    if hay_filtro_dx:
        mask_dx = pd.Series(False, index=df_fev_sub.index)
        cols_fev_dx = ['coddiagnosticoprincipal', 'coddiagnosticorelacionado',
                       'coddiagnosticorelacionado1', 'coddiagnosticorelacionado2']
        for col_fev in cols_fev_dx:
            if col_fev in df_fev_sub.columns:
                for col_regla in cols_dx_regla:
                    if col_regla in regla and not pd.isna(regla[col_regla]) and str(regla[col_regla]).strip() != '':
                        mask_dx |= _evaluar_condicion(df_fev_sub[col_fev], regla[col_regla])
        mask &= mask_dx

    return df_fev_sub[mask]


def _clasificar_curso_vida(edad_meses, edad_dias=-1):
    """
    Clasifica el curso de vida según la edad en días y meses (Resolución 3280 de 2018).
    0 - 7 días: RECIÉN NACIDO
    8 días a 71 meses (5 años): PRIMERA INFANCIA
    """
    if (edad_dias >= 0 and edad_dias <= 7) or (edad_meses == 0 and edad_dias >= 0 and edad_dias <= 7):
        return 'RECIÉN NACIDO'
    elif edad_meses < 0 and edad_dias < 0:
        return 'INDETERMINADO'
    elif edad_meses <= 71:
        return 'PRIMERA INFANCIA'
    elif edad_meses <= 143:
        return 'INFANCIA'
    elif edad_meses <= 215:
        return 'ADOLESCENCIA'
    elif edad_meses <= 347:
        return 'JUVENTUD'
    elif edad_meses <= 707:
        return 'ADULTEZ'
    else:
        return 'VEJEZ'


def _obtener_frecuencia_normativa_texto(freq_meses: int) -> str:
    """
    Convierte el número de meses de frecuencia en un texto claro y descriptivo.
    """
    if freq_meses <= 1:
        return "Cada 1 Mes (Mensual)"
    elif freq_meses <= 3:
        return "Cada 3 Meses (Trimestral)"
    elif freq_meses <= 6:
        return "Cada 6 Meses (Semestral)"
    elif freq_meses <= 12:
        return "Cada 12 Meses (Anual)"
    elif freq_meses <= 24:
        return "Cada 24 Meses (Cada 2 Años)"
    elif freq_meses <= 36:
        return "Cada 36 Meses (Cada 3 Años)"
    elif freq_meses <= 60:
        return "Cada 60 Meses (Cada 5 Años)"
    else:
        return f"Cada {freq_meses} Meses"


def _obtener_rango_edad_norma_texto(edad_min_m: int, edad_max_m: int) -> str:
    """
    Formatea el rango de edad de la regla en términos legibles (Meses/Años).
    """
    if edad_min_m < 24:
        min_str = f"{edad_min_m} Meses"
    else:
        min_str = f"{edad_min_m // 12} Años"

    if edad_max_m >= 1440:
        return f"De {min_str} en adelante ({edad_min_m}+ Meses)"
    elif edad_max_m < 24:
        max_str = f"{edad_max_m} Meses"
    else:
        max_str = f"{edad_max_m // 12} Años"

    return f"De {min_str} a {max_str} ({edad_min_m} a {edad_max_m} Meses)"


def _formatear_edad_paciente(edad_meses, edad_dias=-1):
    """
    Retorna (edad_paciente_texto, unidad_edad).
    - Muestra la edad en DÍAS únicamente cuando el paciente tiene 0 meses. Ej: ("15 Días", "Días")
    - Para pacientes de 1 a 23 meses: Ej: ("13 Meses", "Meses")
    - Para pacientes de 24+ meses: Ej: ("5 Años (60 Meses)", "Años")
    """
    if (pd.isna(edad_meses) or edad_meses < 0) and (pd.isna(edad_dias) or edad_dias < 0):
        return "Edad no registrada", "Sin registro"

    edad_m = int(edad_meses) if not pd.isna(edad_meses) and edad_meses >= 0 else 0

    if edad_m == 0:
        dias = int(edad_dias) if not pd.isna(edad_dias) and edad_dias >= 0 else 0
        return f"{dias} Días", "Días"
    elif edad_m < 24:
        return f"{edad_m} Meses", "Meses"
    else:
        anios = edad_m // 12
        return f"{anios} Años ({edad_m} Meses)", "Años"


def _obtener_frecuencia_segun_edad(actividad: str, id_regla: str, edad_meses: int, edad_dias: int = -1) -> str:
    """
    Determina la frecuencia o esquema normativo específico según la edad del paciente
    basado exactamente en la tabla del Programa de Atención de Salud Infantil (Resolución 3280 de 2018).
    """
    act_norm = actividad.lower().strip()
    id_str = str(id_regla).strip()

    # 1. Recién Nacido (0 a 7 días)
    if (edad_dias >= 0 and edad_dias <= 7) or (edad_meses == 0 and edad_dias >= 0 and edad_dias <= 7):
        if any(kw in act_norm for kw in ['recien nacido', 'recién nacido', 'medicina', 'enfermeria', 'pediatria', 'consulta']):
            return "0-7 Días (Recién Nacido): Control del recién nacido a los 3-5 días de nacido"
        elif 'vacun' in act_norm:
            return "0-7 Días (Recién Nacido): Vacunación de Recién Nacido (BCG, Hepatitis B)"
        return "0-7 Días (Recién Nacido): Atención al Recién Nacido"

    # 2. Primera Infancia (8 días a 71 meses) e Infancia (6 a 11 años / 72 a 143 meses)
    if edad_meses <= 143:

        # --- A. VALORACIÓN INTEGRAL: MEDICINA GENERAL / PEDIATRÍA / MEDICINA FAMILIAR ---
        if any(kw in act_norm for kw in ['medicina general', 'pediatria', 'medicina familiar']) or (id_str in ['1', '4']):
            if edad_meses <= 1:
                return "0-1 Mes: Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif edad_meses in [2, 3]:
                return "2-3 Meses: Le corresponde Consulta por Enfermería (Medicina en 4-5 Meses)"
            elif edad_meses in [4, 5]:
                return "4-5 Meses: Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif edad_meses in [6, 7, 8]:
                return "6-8 Meses: Le corresponde Consulta por Enfermería"
            elif edad_meses in [9, 10, 11]:
                return "9-11 Meses: Le corresponde Consulta por Enfermería"
            elif 12 <= edad_meses <= 17:
                return "12-17 Meses (1 Año): Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 18 <= edad_meses <= 23:
                return "18-23 Meses: Le corresponde Consulta por Enfermería"
            elif 24 <= edad_meses <= 29:
                return "24-29 Meses (2 Años): Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 30 <= edad_meses <= 35:
                return "30-35 Meses: Le corresponde Consulta por Enfermería"
            elif 36 <= edad_meses <= 47:
                return "3 Años (36-47m): Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 48 <= edad_meses <= 59:
                return "4 Años (48-59m): Le corresponde Consulta por Enfermería"
            elif 60 <= edad_meses <= 71:
                return "5 Años (60-71m): Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 72 <= edad_meses <= 83:
                return "6 Años: Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 84 <= edad_meses <= 95:
                return "7 Años: Le corresponde Consulta por Enfermería"
            elif 96 <= edad_meses <= 107:
                return "8 Años: Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 108 <= edad_meses <= 119:
                return "9 Años: Le corresponde Consulta por Enfermería"
            elif 120 <= edad_meses <= 131:
                return "10 Años: Consulta Valoración Integral Medicina General / Pediatría (X)"
            elif 132 <= edad_meses <= 143:
                return "11 Años: Le corresponde Consulta por Enfermería"

        # --- B. VALORACIÓN INTEGRAL: PROFESIONAL DE ENFERMERÍA ---
        if 'enfermeria' in act_norm or (id_str in ['2', '5']):
            if edad_meses <= 1:
                return "0-1 Mes: Le corresponde Consulta Medicina General / Pediatría"
            elif edad_meses in [2, 3]:
                return "2-3 Meses: Consulta Valoración Integral por Enfermería (X)"
            elif edad_meses in [4, 5]:
                return "4-5 Meses: Le corresponde Consulta Medicina General / Pediatría"
            elif edad_meses in [6, 7, 8]:
                return "6-8 Meses: Consulta Valoración Integral por Enfermería (X)"
            elif edad_meses in [9, 10, 11]:
                return "9-11 Meses: Consulta Valoración Integral por Enfermería (X)"
            elif 12 <= edad_meses <= 17:
                return "12-17 Meses: Le corresponde Consulta Medicina General / Pediatría"
            elif 18 <= edad_meses <= 23:
                return "18-23 Meses: Consulta Valoración Integral por Enfermería (X)"
            elif 24 <= edad_meses <= 29:
                return "24-29 Meses: Le corresponde Consulta Medicina General / Pediatría"
            elif 30 <= edad_meses <= 35:
                return "30-35 Meses: Consulta Valoración Integral por Enfermería (X)"
            elif 36 <= edad_meses <= 47:
                return "3 Años: Le corresponde Consulta Medicina General / Pediatría"
            elif 48 <= edad_meses <= 59:
                return "4 Años: Consulta Valoración Integral por Enfermería (X)"
            elif 60 <= edad_meses <= 71:
                return "5 Años: Le corresponde Consulta Medicina General / Pediatría"
            elif 72 <= edad_meses <= 83:
                return "6 Años: Le corresponde Consulta Medicina General / Pediatría"
            elif 84 <= edad_meses <= 95:
                return "7 Años: Consulta Valoración Integral por Enfermería (X)"
            elif 96 <= edad_meses <= 107:
                return "8 Años: Le corresponde Consulta Medicina General / Pediatría"
            elif 108 <= edad_meses <= 119:
                return "9 Años: Consulta Valoración Integral por Enfermería (X)"
            elif 120 <= edad_meses <= 131:
                return "10 Años: Le corresponde Consulta Medicina General / Pediatría"
            elif 132 <= edad_meses <= 143:
                return "11 Años: Consulta Valoración Integral por Enfermería (X)"

        # --- C. ODONTOLOGÍA / SALUD BUCAL ---
        if any(kw in act_norm for kw in ['odontologia', 'salud bucal']) or (id_str == '11'):
            if edad_meses < 6:
                return "0-5 Meses: Consulta de Salud Bucal a partir de los 6 meses de edad"
            else:
                return "A partir de los 6 meses: Consulta de Salud Bucal Una vez al año"

        # --- D. PROFILAXIS Y REMOCIÓN DE PLACA BACTERIANA ---
        if any(kw in act_norm for kw in ['profilaxis', 'placa']) or (id_str in ['12', '14']):
            if edad_meses < 12:
                return "0-11 Meses: Profilaxis a partir del año de edad"
            else:
                return "A partir del año de edad: Profilaxis y Remoción de Placa Semestral (2 veces al año)"

        # --- E. APLICACIÓN DE BARNIZ DE FLÚOR ---
        if 'fluor' in act_norm or (id_str == '13'):
            if edad_meses < 12:
                return "0-11 Meses: Barniz de Flúor a partir del año de edad"
            else:
                return "A partir del año de edad: Aplicación de Barniz de Flúor Semestral (2 veces al año)"

        # --- F. APLICACIÓN DE SELLANTES ---
        if 'sellante' in act_norm:
            if edad_meses < 36:
                return "0-2 Años: Aplicación de Sellantes a partir de los 3 años"
            else:
                return "A partir de los 3 años: Según criterio del profesional y control de permanencia"

        # --- G. LACTANCIA MATERNA ---
        if 'lactancia' in act_norm:
            if edad_meses <= 1:
                return "0-1 Mes: Atención y apoyo a la Lactancia Materna (X)"
            else:
                return "2-3 Meses+: Según hallazgos y criterio del profesional"

        # --- H. ANEMIA / HEMOGLOBINA ---
        if any(kw in act_norm for kw in ['anemia', 'hemoglobina']) or (id_str == '24'):
            if 120 <= edad_meses <= 143:
                return "10 a 13 Años: Tamizaje para anemia Una vez entre los 10 y 13 años"
            else:
                return "1 a 9 Años: Tamizaje para anemia Una vez según riesgo identificado"

        # --- I. VACUNACIÓN ---
        if 'vacun' in act_norm:
            if edad_meses in [2, 3]:
                return "2-3 Meses: Esquema de Vacunación (2 meses)"
            elif edad_meses in [4, 5]:
                return "4-5 Meses: Esquema de Vacunación (4 meses)"
            elif edad_meses in [6, 7, 8]:
                return "6-8 Meses: Esquema de Vacunación (6 meses)"
            elif 12 <= edad_meses <= 17:
                return "12-17 Meses: Esquema de Vacunación (12 meses)"
            elif 18 <= edad_meses <= 23:
                return "18-23 Meses: Refuerzo de Vacunación (18 meses)"
            elif 60 <= edad_meses <= 71:
                return "5 Años: Refuerzo de Vacunación (5 años)"
            elif 108 <= edad_meses <= 119:
                return "9 Años: Vacunación contra VPH"
            return "Según Esquema PAI de Vacunación"

    # 3. Adolescencia (12 a 17 años / 144 - 215 meses)
    if 144 <= edad_meses <= 215:
        return "Adolescencia (12-17 Años): Atención cada 12 meses (Anual)"

    # 4. Juventud (18 a 28 años / 216 - 347 meses)
    if 216 <= edad_meses <= 347:
        return "Juventud (18-28 Años): Atención cada 12 meses (Anual)"

    # 5. Adultez (29 a 59 años / 348 - 707 meses)
    if 348 <= edad_meses <= 707:
        if any(kw in act_norm for kw in ['citologia', 'vph']):
            return "25-65 Años: Esquema Citología 1-1-3 o ADN-VPH cada 5 Años"
        elif 'mamografia' in act_norm:
            return "50-69 Años: Mamografía cada 2 Años"
        return "Adultez (29-59 Años): Atención cada 12 meses (Anual)"

    # 6. Vejez (60+ años / 708+ meses)
    if edad_meses >= 708:
        return "Vejez (60+ Años): Atención cada 6 meses (Semestral)"

    return "Según Esquema de Resolución 3280"

def _aplica_en_este_momento(actividad: str, id_regla: str, edad_meses: int, edad_dias: int = -1) -> bool:
    """
    Filtro estricto (Resolución 3280): Determina si la actividad aplica EXACTAMENTE 
    en este mes de vida para niños. Evita que rangos amplios como 0-71 meses asusten con pendientes irreales.
    """
    act_norm = actividad.lower().strip()
    id_str = str(id_regla).strip()
    
    if pd.isna(edad_meses) or edad_meses < 0:
        return True # Si no hay edad, dejamos que el rango general actúe

    if edad_meses <= 143:
        # Medicina General / Pediatría
        if any(kw in act_norm for kw in ['medicina general', 'pediatria', 'medicina familiar']) or (id_str in ['1', '3', '4', '6']):
            # Meses exactos donde toca Medicina General: 1, 4-5, 12-17, 24-29, 36-47, 60-71, 72-83, 96-107, 120-131
            meses_val = [0, 1, 4, 5] + list(range(12, 18)) + list(range(24, 30)) + list(range(36, 48)) + list(range(60, 72)) + list(range(72, 84)) + list(range(96, 108)) + list(range(120, 132))
            if edad_meses not in meses_val:
                return False

        # Enfermería
        elif 'enfermeria' in act_norm or (id_str in ['2', '5', '7']):
            # Meses exactos: 2-3, 6-8, 9-11, 18-23, 30-35, 48-59, 84-95, 108-119, 132-143
            meses_val = [2, 3, 6, 7, 8, 9, 10, 11] + list(range(18, 24)) + list(range(30, 36)) + list(range(48, 60)) + list(range(84, 96)) + list(range(108, 120)) + list(range(132, 144))
            if edad_meses not in meses_val:
                return False
                
        # Odontología
        elif any(kw in act_norm for kw in ['odontologia', 'salud bucal']) or (id_str == '11'):
            if edad_meses < 6: return False
            
        # Profilaxis, Placa, Flúor
        elif any(kw in act_norm for kw in ['profilaxis', 'placa', 'fluor']) or (id_str in ['12', '13', '14']):
            if edad_meses < 12: return False

    return True


def ejecutar_motor_reglas(df_pacientes: pd.DataFrame, df_fev: pd.DataFrame, docs_cohorte: set = None, df_nominal: pd.DataFrame = None) -> tuple:
    """
    Motor de reglas con evaluación de frecuencia según Resolución 3280 de 2018.
    
    FASE 1: Determina qué actividades necesita cada paciente según sus datos
            demográficos (sexo, edad) y las reglas del catálogo Excel.
    
    FASE 2: Verifica en la FEV cuáles de esas actividades ya fueron realizadas,
            evaluando los criterios clínicos (CUPS, finalidad, diagnósticos).
            Evalúa si la atención más reciente está VIGENTE o VENCIDA según
            la frecuencia normativa de cada actividad.
    
    Parámetros:
        df_pacientes: Pacientes a gestionar (Población IPS con edad_meses).
        df_fev:       Atenciones FEV/RIPS completas.
    
    Retorna:
        (df_necesarias, df_realizadas, df_pendientes)
    """
    separador("EJECUTANDO MOTOR DE REGLAS")

    cols_realizadas = ['nro_identificacion', 'actividad', 'columna_nominal', 'fecha_atencion',
                       'id_regla', 'frecuencia_meses', 'vigente']
    cols_pendientes = ['nro_identificacion', 'actividad', 'columna_nominal', 'id_regla',
                       'frecuencia_meses', 'curso_vida']

    if df_pacientes.empty:
        print(" [ADVERTENCIA] No hay pacientes a gestionar.")
        return pd.DataFrame(columns=cols_pendientes), pd.DataFrame(columns=cols_realizadas), pd.DataFrame(columns=cols_pendientes)

    # Cargar reglas activas
    df_reglas = cargar_reglas()
    if df_reglas.empty:
        return pd.DataFrame(columns=cols_pendientes), pd.DataFrame(columns=cols_realizadas), pd.DataFrame(columns=cols_pendientes)

    # Pre-filtrar FEV: solo registros de pacientes que estamos gestionando
    col_doc = 'nro_identificacion'
    docs_gestionar = set(df_pacientes[col_doc].astype(str).str.strip().str.upper())

    if df_fev is not None and not df_fev.empty and col_doc in df_fev.columns:
        df_fev_filtrada = df_fev[
            df_fev[col_doc].astype(str).str.strip().str.upper().isin(docs_gestionar)
        ].copy()
        print(f" -> FEV filtrada a pacientes de la población: {len(df_fev_filtrada):,} de {len(df_fev):,} registros.")
    else:
        df_fev_filtrada = pd.DataFrame()
        print(" [ADVERTENCIA] No hay datos FEV disponibles para verificar actividades.")

    todas_necesarias = []
    todas_realizadas = []

    print(f" -> Evaluando {len(df_pacientes):,} pacientes contra {len(df_reglas)} reglas activas...\n")

    # Fecha de corte para evaluar vigencia
    fecha_corte = pd.Timestamp.now()

    # =========================================================================
    # FASE 1 + 2: Para cada regla, determinar pacientes que la necesitan
    #             y verificar si tienen atención en FEV
    # =========================================================================
    for _, regla in df_reglas.iterrows():
        id_regla = str(regla.get('ID', 'SIN_ID')).strip()
        actividad = str(regla.get('ACTIVIDAD', '')).strip()
        columna_nominal = str(regla.get('COLUMNA_NOMINAL', '')).strip().lower()

        # Frecuencia normativa (Resolución 3280)
        try:
            frecuencia_meses = int(float(str(regla.get('FRECUENCIA_MESES', '12')).strip()))
        except (ValueError, TypeError):
            frecuencia_meses = 12

        # --- FASE 1: ¿Qué pacientes necesitan esta actividad? (Filtros demográficos) ---
        mask_pac = pd.Series(True, index=df_pacientes.index)

        # Filtro: Sexo
        if 'SEXO' in regla and not pd.isna(regla['SEXO']) and str(regla['SEXO']).strip() != '':
            if 'sexo' in df_pacientes.columns:
                sexo_regla = str(regla['SEXO']).strip().upper()
                mask_pac &= (df_pacientes['sexo'].astype(str).str.strip().str.upper() == sexo_regla)

        # Filtro: Edad Mínima (en meses)
        edad_min = 0
        if 'EDAD_MIN_MESES' in regla and not pd.isna(regla['EDAD_MIN_MESES']) and str(regla['EDAD_MIN_MESES']).strip() != '':
            try:
                edad_min = int(float(str(regla['EDAD_MIN_MESES']).strip()))
                mask_pac &= (df_pacientes['edad_meses'] >= edad_min)
            except ValueError:
                pass

        # Filtro: Edad Máxima (en meses)
        edad_max = 1440
        if 'EDAD_MAX_MESES' in regla and not pd.isna(regla['EDAD_MAX_MESES']) and str(regla['EDAD_MAX_MESES']).strip() != '':
            try:
                edad_max = int(float(str(regla['EDAD_MAX_MESES']).strip()))
                mask_pac &= (df_pacientes['edad_meses'] <= edad_max)
            except ValueError:
                pass

        # Filtro estricto (mes a mes) Res 3280
        mask_estricta = [
            _aplica_en_este_momento(actividad, id_regla, em, ed)
            for em, ed in zip(df_pacientes['edad_meses'], df_pacientes.get('edad_dias', [-1]*len(df_pacientes)))
        ]
        mask_pac &= pd.Series(mask_estricta, index=df_pacientes.index)

        pacientes_necesitan = df_pacientes[mask_pac]

        if pacientes_necesitan.empty:
            print(f"    Regla {id_regla} ({actividad}) [cada {frecuencia_meses}m]: 0 pacientes. Saltando.")
            continue

        # Clasificar curso de vida y texto legible de norma/frecuencia
        dias_vals = pacientes_necesitan['edad_dias'].values if 'edad_dias' in pacientes_necesitan.columns else [-1] * len(pacientes_necesitan)
        
        # === OPTIMIZACION: FILTRAR A COHORTE PARA 'NECESARIAS' ===
        if docs_cohorte and len(docs_cohorte) > 0:
            mask_cohorte = pacientes_necesitan[col_doc].astype(str).str.strip().str.upper().isin(docs_cohorte)
            pacientes_necesitan_cohorte = pacientes_necesitan[mask_cohorte].copy()
            dias_vals_cohorte = dias_vals[mask_cohorte]
        else:
            pacientes_necesitan_cohorte = pacientes_necesitan.copy()
            dias_vals_cohorte = dias_vals

        cursos_vida = [
            _clasificar_curso_vida(m, d)
            for m, d in zip(pacientes_necesitan_cohorte['edad_meses'].values, dias_vals_cohorte)
        ]
        frecuencia_texto = _obtener_frecuencia_normativa_texto(frecuencia_meses)
        rango_norma_texto = _obtener_rango_edad_norma_texto(edad_min, edad_max)

        # Registrar las actividades necesarias con frecuencia y curso de vida (SOLO COHORTE)
        if not pacientes_necesitan_cohorte.empty:
            df_necesaria = pd.DataFrame({
                'nro_identificacion': pacientes_necesitan_cohorte[col_doc].values,
                'actividad': actividad,
                'columna_nominal': columna_nominal,
                'id_regla': id_regla,
                'frecuencia_meses': frecuencia_meses,
                'frecuencia_normativa': frecuencia_texto,
                'rango_edad_norma': rango_norma_texto,
                'curso_vida': cursos_vida
            })
            todas_necesarias.append(df_necesaria)

        # --- FASE 2: ¿Cuáles de esos pacientes tienen atención en FEV? ---
        if df_fev_filtrada.empty:
            print(f"    Regla {id_regla} ({actividad}) [cada {frecuencia_meses}m]: {len(pacientes_necesitan):,} la necesitan | 0 realizadas (sin FEV)")
            continue

        # Filtrar FEV a solo los pacientes que necesitan esta actividad
        docs_necesitan = set(pacientes_necesitan[col_doc].astype(str).str.strip().str.upper())
        df_fev_sub = df_fev_filtrada[
            df_fev_filtrada[col_doc].astype(str).str.strip().str.upper().isin(docs_necesitan)
        ]

        if df_fev_sub.empty:
            print(f"    Regla {id_regla} ({actividad}) [cada {frecuencia_meses}m]: {len(pacientes_necesitan):,} la necesitan | 0 realizadas")
            continue

        # Aplicar filtros clínicos de la regla sobre FEV
        atenciones_match = _filtrar_fev_por_regla(df_fev_sub, regla)

        if not atenciones_match.empty:
            # Obtener columna de fecha de la regla
            campo_fecha = str(regla.get('CAMPO_FECHA', 'fechainicioatencion')).strip().lower()
            col_fecha = campo_fecha if campo_fecha in atenciones_match.columns else 'fechainicioatencion'

            fechas_atencion = pd.to_datetime(atenciones_match[col_fecha], errors='coerce')

            # Evaluar vigencia: ¿la atención está dentro de la ventana de frecuencia?
            fecha_limite = fecha_corte - pd.DateOffset(months=frecuencia_meses)
            vigencia = fechas_atencion >= fecha_limite

            df_real = pd.DataFrame({
                'nro_identificacion': atenciones_match[col_doc].values,
                'actividad': actividad,
                'columna_nominal': columna_nominal,
                'fecha_atencion': fechas_atencion.values,
                'id_regla': id_regla,
                'frecuencia_meses': frecuencia_meses,
                'vigente': vigencia.values
            })
            todas_realizadas.append(df_real)

        n_realizadas = len(atenciones_match[col_doc].unique()) if not atenciones_match.empty else 0
        n_vigentes = int(vigencia.sum()) if not atenciones_match.empty else 0
        print(f"    Regla {id_regla} ({actividad}) [cada {frecuencia_meses}m]: {len(pacientes_necesitan):,} la necesitan | {n_realizadas:,} con FEV ({n_vigentes} vigentes)")

    # =========================================================================
    # CONSOLIDAR RESULTADOS
    # =========================================================================
    # Actividades necesarias
    if todas_necesarias:
        df_necesarias = pd.concat(todas_necesarias, ignore_index=True)
    else:
        df_necesarias = pd.DataFrame(columns=cols_pendientes)

    # Actividades realizadas (quedarse con la más reciente por paciente × regla)
    if todas_realizadas:
        df_realizadas = pd.concat(todas_realizadas, ignore_index=True)
        df_realizadas['fecha_atencion'] = pd.to_datetime(df_realizadas['fecha_atencion'], errors='coerce')
        df_realizadas = df_realizadas.sort_values('fecha_atencion', ascending=False)
        df_realizadas = df_realizadas.drop_duplicates(subset=['nro_identificacion', 'id_regla'], keep='first')
    else:
        df_realizadas = pd.DataFrame(columns=cols_realizadas)

    # Actividades pendientes: necesarias SIN atención vigente en FEV NI en la Nominal
    # Un paciente está PENDIENTE si:
    #   a) Nunca tuvo la atención (ni en FEV ni en Nominal), o
    #   b) Su atención más reciente está VENCIDA (fuera de la ventana de frecuencia)
    if not df_necesarias.empty:
        df_pendientes = df_necesarias.copy()
        df_pendientes['_doc_norm'] = df_pendientes['nro_identificacion'].astype(str).str.strip().str.upper()

        # 1. Descartar los que ya están vigentes en FEV
        if not df_realizadas.empty:
            realizadas_vigentes = df_realizadas[df_realizadas['vigente'] == True][['nro_identificacion', 'id_regla']].copy()
            realizadas_vigentes['nro_identificacion'] = realizadas_vigentes['nro_identificacion'].astype(str).str.strip().str.upper()
            realizadas_vigentes = realizadas_vigentes.drop_duplicates()
            realizadas_vigentes['_vigente_fev'] = True

            df_pendientes = df_pendientes.merge(
                realizadas_vigentes,
                left_on=['_doc_norm', 'id_regla'],
                right_on=['nro_identificacion', 'id_regla'],
                how='left',
                suffixes=('', '_r')
            )
            df_pendientes = df_pendientes[df_pendientes['_vigente_fev'].isna()].copy()
            cols_drop = [c for c in ['_vigente_fev', 'nro_identificacion_r'] if c in df_pendientes.columns]
            df_pendientes = df_pendientes.drop(columns=cols_drop)

        # 2. Descartar los que ya están vigentes en la Nominal (EPS)
        if df_nominal is not None and not df_nominal.empty:
            # Preparar nominal para cruce
            df_nom_cruce = df_nominal.copy()
            df_nom_cruce['_doc_norm'] = df_nom_cruce['nro_identificacion'].astype(str).str.strip().str.upper()
            df_nom_cruce = df_nom_cruce.drop_duplicates(subset=['_doc_norm'])

            # Unir columnas de la nominal a pendientes
            df_pendientes = df_pendientes.merge(df_nom_cruce, on='_doc_norm', how='left', suffixes=('', '_nom'))

            # Evaluar vigencia de cada registro según su columna_nominal
            mascara_pendiente = pd.Series(True, index=df_pendientes.index)
            
            for idx, row in df_pendientes.iterrows():
                col_nom = str(row.get('columna_nominal', '')).strip()
                if col_nom and col_nom in df_pendientes.columns:
                    fecha_nominal = row.get(col_nom)
                    if not pd.isna(fecha_nominal):
                        try:
                            fecha_nom_dt = pd.to_datetime(fecha_nominal)
                            freq_meses = row.get('frecuencia_meses', 12)
                            fecha_limite = fecha_corte - pd.DateOffset(months=freq_meses)
                            if fecha_nom_dt >= fecha_limite:
                                # Está vigente en la nominal, NO es pendiente
                                mascara_pendiente[idx] = False
                        except Exception:
                            pass # Si la fecha es inválida, sigue siendo pendiente
            
            df_pendientes = df_pendientes[mascara_pendiente].copy()
            
            # Limpiar columnas extras traídas de la nominal
            cols_to_keep = list(df_necesarias.columns) + ['_doc_norm']
            df_pendientes = df_pendientes[[c for c in cols_to_keep if c in df_pendientes.columns]].copy()

        df_pendientes = df_pendientes.drop(columns=['_doc_norm'], errors='ignore')
    else:
        df_pendientes = df_necesarias.copy()

    # Enrich df_pendientes con campos demográficos, de contacto y de explicación de la norma
    if not df_pendientes.empty and not df_pacientes.empty:
        cols_info_paciente = [
            'nro_identificacion', 'eapb', 'tipo_identificacion', 'nombre_completo',
            'sexo', 'edad_actual', 'edad_meses', 'edad_dias', 'curso_vida', 'celular', 'celular2',
            'telefono_fijo', 'direccion_residencia', 'barrio', 'comuna',
            'correo_electronico', 'nombre_ips'
        ]
        cols_existentes = [c for c in cols_info_paciente if c in df_pacientes.columns]
        df_pac_info = df_pacientes[cols_existentes].drop_duplicates(subset=['nro_identificacion']).copy()

        df_pendientes = df_pendientes.merge(
            df_pac_info,
            on='nro_identificacion',
            how='left',
            suffixes=('', '_pob')
        )

        # Generar edad legible del paciente, unidad y explicación clara
        edad_textos = []
        unidades_edad = []
        frecuencias_edad = []
        detalles_debe = []

        for _, row in df_pendientes.iterrows():
            e_meses = row.get('edad_meses', -1)
            e_dias = row.get('edad_dias', -1)
            e_texto, u_edad = _formatear_edad_paciente(e_meses, e_dias)
            c_vida = str(row.get('curso_vida', '')).strip()
            act = str(row.get('actividad', '')).strip()
            id_r = str(row.get('id_regla', '')).strip()
            r_norma = str(row.get('rango_edad_norma', '')).strip()
            f_norma = str(row.get('frecuencia_normativa', '')).strip()

            freq_edad = _obtener_frecuencia_segun_edad(act, id_r, e_meses, e_dias)

            detalle = (
                f"Paciente de {e_texto} (Curso de Vida: {c_vida}). "
                f"Debe la atención '{act}'. Frecuencia según edad: [{freq_edad}]. "
                f"Aplica para rango {r_norma} con frecuencia obligatoria base de {f_norma} (Res. 3280)."
            )

            edad_textos.append(e_texto)
            unidades_edad.append(u_edad)
            frecuencias_edad.append(freq_edad)
            detalles_debe.append(detalle)

        df_pendientes['edad_paciente'] = edad_textos
        df_pendientes['unidad_edad'] = unidades_edad
        df_pendientes['frecuencia_segun_edad'] = frecuencias_edad
        df_pendientes['detalle_atencion_debe'] = detalles_debe

        # Columnas de seguimiento operativo para Call Center/Gestores
        df_pendientes['estado_gestion'] = ""
        df_pendientes['fecha_cita_programada'] = ""
        df_pendientes['observaciones'] = ""

        # Ordenar columnas de forma limpia y lógica
        cols_ordenadas = [
            'eapb', 'tipo_identificacion', 'nro_identificacion', 'nombre_completo',
            'sexo', 'edad_paciente', 'unidad_edad', 'curso_vida',
            'actividad', 'frecuencia_segun_edad', 'rango_edad_norma', 'detalle_atencion_debe',
            'celular', 'celular2', 'telefono_fijo', 'direccion_residencia',
            'barrio', 'comuna', 'correo_electronico', 'nombre_ips',
            'id_regla',
            'estado_gestion', 'fecha_cita_programada', 'observaciones'
        ]
        cols_finales = [c for c in cols_ordenadas if c in df_pendientes.columns]
        df_pendientes = df_pendientes[cols_finales].copy()

    # Resumen final
    n_vigentes_total = int(df_realizadas['vigente'].sum()) if not df_realizadas.empty and 'vigente' in df_realizadas.columns else 0
    n_vencidas_total = len(df_realizadas) - n_vigentes_total if not df_realizadas.empty else 0

    separador("RESUMEN DEL MOTOR DE REGLAS (POBLACIÓN IPS)")
    print(f" -> Actividades necesarias (total)  : {len(df_necesarias):,}")
    print(f" -> Atenciones encontradas en FEV   : {len(df_realizadas):,}")
    print(f"    |-- VIGENTES (dentro de frecuencia) : {n_vigentes_total:,}")
    print(f"    +-- VENCIDAS (fuera de frecuencia)  : {n_vencidas_total:,}")
    print(f" -> Actividades PENDIENTES          : {len(df_pendientes):,}")
    if len(df_necesarias) > 0:
        pct = (n_vigentes_total / len(df_necesarias)) * 100
        print(f" -> Cumplimiento vigente            : {pct:.1f}%")

    return df_necesarias, df_realizadas, df_pendientes