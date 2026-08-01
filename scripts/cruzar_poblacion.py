# -*- coding: utf-8 -*-
"""
Módulo: cruzar_poblacion.py
Descripción: Cruza la cohorte nominal (pacientes de cada EAPB) con la población
             asignada a la IPS para determinar qué pacientes se deben gestionar.
             Enriquece los datos del paciente con información demográfica y de contacto.
"""

import pandas as pd
import numpy as np
from utilidades import separador


def _construir_nombre_completo(df: pd.DataFrame) -> pd.Series:
    """
    Combina primer_nombre, segundo_nombre, primer_apellido y segundo_apellido en una sola cadena limpia.
    """
    p_nom = df.get('primer_nombre', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    s_nom = df.get('segundo_nombre', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    p_ape = df.get('primer_apellido', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()
    s_ape = df.get('segundo_apellido', pd.Series('', index=df.index)).fillna('').astype(str).str.strip()

    nombre_completo = (p_nom + ' ' + s_nom).str.strip() + ' ' + (p_ape + ' ' + s_ape).str.strip()
    return nombre_completo.str.strip().str.upper()


def cruzar_cohorte_poblacion(df_nominal: pd.DataFrame, df_poblacion: pd.DataFrame) -> tuple:
    """
    Cruza los pacientes de la cohorte nominal con la población asignada a la IPS.

    Retorna:
        - df_pacientes: DataFrame con los pacientes que están en AMBAS bases (a gestionar),
                        enriquecidos con datos demográficos, curso de vida y de contacto.
        - docs_solo_cohorte: set de documentos que están en cohorte pero NO en población.
        - docs_solo_poblacion: set de documentos que están en población pero NO en cohorte.
    """
    separador("CRUCE: COHORTE × POBLACIÓN ASIGNADA")

    col_doc = 'nro_identificacion'

    if col_doc not in df_nominal.columns:
        print(f" [ERROR CRÍTICO] La columna '{col_doc}' no existe en la nominal.")
        return pd.DataFrame(), set(), set()
    if col_doc not in df_poblacion.columns:
        print(f" [ERROR CRÍTICO] La columna '{col_doc}' no existe en la población.")
        return pd.DataFrame(), set(), set()

    # Normalizar documentos para comparación limpia
    docs_nominal = set(
        df_nominal[col_doc].astype(str).str.strip().str.upper()
        .loc[lambda s: ~s.isin(['', 'NAN', 'NONE', 'NAT'])]
    )
    docs_poblacion = set(
        df_poblacion[col_doc].astype(str).str.strip().str.upper()
        .loc[lambda s: ~s.isin(['', 'NAN', 'NONE', 'NAT'])]
    )

    docs_comunes = docs_nominal & docs_poblacion
    docs_solo_cohorte = docs_nominal - docs_poblacion
    docs_solo_poblacion = docs_poblacion - docs_nominal

    print(f" -> Pacientes en Cohorte (Nominal)     : {len(docs_nominal):,}")
    print(f" -> Pacientes en Población (BD IPS)    : {len(docs_poblacion):,}")
    print(f" -> Pacientes en AMBAS bases (a gestionar) : {len(docs_comunes):,}")
    if docs_solo_cohorte:
        print(f" -> Solo en Cohorte (sin población)     : {len(docs_solo_cohorte):,}")
    if docs_solo_poblacion:
        print(f" -> Solo en Población (sin cohorte)     : {len(docs_solo_poblacion):,}")

    # Filtrar la nominal para los pacientes del cruce
    df_pacientes = df_nominal[
        df_nominal[col_doc].astype(str).str.strip().str.upper().isin(docs_comunes)
    ].copy()

    # Crear clave limpia para merge con población
    df_pacientes['_doc_norm'] = df_pacientes[col_doc].astype(str).str.strip().str.upper()

    # Preparar campos de contacto de la población
    cols_contacto_pob = [
        col_doc, 'celular', 'celular2', 'telefono_fijo',
        'direccion_residencia', 'barrio', 'comuna',
        'correo_electronico', 'ciclovida', 'nombreips'
    ]
    cols_existentes_pob = [c for c in cols_contacto_pob if c in df_poblacion.columns]

    if len(cols_existentes_pob) > 1:
        df_pob_contacto = df_poblacion[cols_existentes_pob].copy()
        df_pob_contacto['_doc_norm'] = df_pob_contacto[col_doc].astype(str).str.strip().str.upper()
        df_pob_contacto = df_pob_contacto.drop(columns=[col_doc]).drop_duplicates(subset=['_doc_norm'])

        # Combinar datos de contacto de población en df_pacientes
        df_pacientes = df_pacientes.merge(
            df_pob_contacto,
            on='_doc_norm',
            how='left',
            suffixes=('', '_pob')
        )

    # Construir nombre completo
    df_pacientes['nombre_completo'] = _construir_nombre_completo(df_pacientes)

    # Unificar curso de vida (usar ciclo_vida o curso_vida si existe)
    if 'curso_vida' not in df_pacientes.columns and 'ciclovida' in df_pacientes.columns:
        df_pacientes['curso_vida'] = df_pacientes['ciclovida']
    elif 'curso_vida' in df_pacientes.columns and 'ciclovida' in df_pacientes.columns:
        df_pacientes['curso_vida'] = df_pacientes['curso_vida'].fillna(df_pacientes['ciclovida'])

    # Calcular edad en meses y días
    if 'fecha_nacimiento' in df_pacientes.columns:
        hoy = pd.Timestamp.now()
        nacimiento = pd.to_datetime(df_pacientes['fecha_nacimiento'], errors='coerce', dayfirst=True)

        edad_meses = (hoy.year - nacimiento.dt.year) * 12 + (hoy.month - nacimiento.dt.month)
        dias_insuficientes = hoy.day < nacimiento.dt.day
        edad_meses = np.where(dias_insuficientes, edad_meses - 1, edad_meses)

        edad_dias = (hoy - nacimiento).dt.days

        df_pacientes['edad_meses'] = pd.Series(edad_meses, index=df_pacientes.index).fillna(-1).astype(int)
        df_pacientes['edad_dias'] = pd.Series(edad_dias, index=df_pacientes.index).fillna(-1).astype(int)
    else:
        print(" [ADVERTENCIA] No se encontró 'fecha_nacimiento' para calcular edad en meses y días.")
        df_pacientes['edad_meses'] = -1
        df_pacientes['edad_dias'] = -1

    if '_doc_norm' in df_pacientes.columns:
        df_pacientes = df_pacientes.drop(columns=['_doc_norm'])

    print(f"\n -> Total pacientes a gestionar: {len(df_pacientes):,}")

    return df_pacientes, docs_solo_cohorte, docs_solo_poblacion


def clasificar_atenciones_cohorte(df_necesarias: pd.DataFrame,
                                   df_realizadas: pd.DataFrame,
                                   df_pendientes: pd.DataFrame,
                                   df_nominal: pd.DataFrame,
                                   df_poblacion: pd.DataFrame) -> tuple:
    """
    Cruza los resultados del motor de reglas (ejecutado sobre la Población IPS x FEV)
    con la Cohorte Nominal (EAPB).

    Clasifica las atenciones en:
      1. Realizadas en Cohorte (se usan para actualizar la nominal).
      2. Realizadas Fuera de Cohorte (atenciones "de más" a pacientes no incluidos en la cohorte EAPB).
      3. Pendientes en Cohorte (pacientes de la cohorte que requieren gestión).

    Retorna:
      (df_realizadas_cohorte, df_realizadas_fuera_cohorte, df_pendientes_cohorte)
    """
    separador("CLASIFICANDO ATENCIONES: COHORTE VS. FUERA DE COHORTE")

    col_doc = 'nro_identificacion'

    if df_nominal is not None and not df_nominal.empty and col_doc in df_nominal.columns:
        docs_cohorte = set(
            df_nominal[col_doc].astype(str).str.strip().str.upper()
            .loc[lambda s: ~s.isin(['', 'NAN', 'NONE', 'NAT'])]
        )
    else:
        docs_cohorte = set()

    print(f" -> Total documentos únicos en Cohorte Nominal: {len(docs_cohorte):,}")

    # =========================================================================
    # 1. CLASIFICAR ATENCIONES REALIZADAS
    # =========================================================================
    if df_realizadas is not None and not df_realizadas.empty:
        df_real = df_realizadas.copy()
        df_real['_doc_norm'] = df_real[col_doc].astype(str).str.strip().str.upper()
        df_real['en_cohorte'] = df_real['_doc_norm'].isin(docs_cohorte)

        df_realizadas_cohorte = df_real[df_real['en_cohorte']].drop(columns=['_doc_norm', 'en_cohorte']).copy()
        df_realizadas_fuera_cohorte = df_real[~df_real['en_cohorte']].drop(columns=['_doc_norm', 'en_cohorte']).copy()
    else:
        df_realizadas_cohorte = pd.DataFrame()
        df_realizadas_fuera_cohorte = pd.DataFrame()

    # =========================================================================
    # 2. CLASIFICAR ATENCIONES PENDIENTES
    # =========================================================================
    if df_pendientes is not None and not df_pendientes.empty:
        df_pend = df_pendientes.copy()
        df_pend['_doc_norm'] = df_pend[col_doc].astype(str).str.strip().str.upper()
        df_pend['en_cohorte'] = df_pend['_doc_norm'].isin(docs_cohorte)

        df_pendientes_cohorte = df_pend[df_pend['en_cohorte']].drop(columns=['_doc_norm', 'en_cohorte']).copy()
    else:
        df_pendientes_cohorte = pd.DataFrame()

    # =========================================================================
    # 3. ENRIQUECER ATENCIONES REALIZADAS CON DATOS DE CONTACTO (COHORTE Y FUERA)
    # =========================================================================
    if df_poblacion is not None and not df_poblacion.empty:
        cols_info_paciente = [
            col_doc, 'eapb', 'tipo_identificacion', 'nombre_completo',
            'sexo', 'edad_actual', 'curso_vida', 'nombre_ips'
        ]
        cols_existentes = [c for c in cols_info_paciente if c in df_poblacion.columns]
        df_pob_info = df_poblacion[cols_existentes].copy()
        df_pob_info[col_doc] = df_pob_info[col_doc].astype(str).str.strip().str.upper()
        df_pob_info = df_pob_info.drop_duplicates(subset=[col_doc])

        cols_ordenadas = [
            'eapb', 'tipo_identificacion', col_doc, 'nombre_completo',
            'sexo', 'edad_actual', 'curso_vida', 'actividad', 'fecha_atencion',
            'nombre_ips',
            'columna_nominal', 'id_regla', 'frecuencia_meses'
        ]

        if not df_realizadas_cohorte.empty:
            df_realizadas_cohorte[col_doc] = df_realizadas_cohorte[col_doc].astype(str).str.strip().str.upper()
            df_realizadas_cohorte = df_realizadas_cohorte.merge(df_pob_info, on=col_doc, how='left')
            cols_finales_c = [c for c in cols_ordenadas if c in df_realizadas_cohorte.columns]
            df_realizadas_cohorte = df_realizadas_cohorte[cols_finales_c].copy()

        if not df_realizadas_fuera_cohorte.empty:
            df_realizadas_fuera_cohorte[col_doc] = df_realizadas_fuera_cohorte[col_doc].astype(str).str.strip().str.upper()
            df_realizadas_fuera_cohorte = df_realizadas_fuera_cohorte.merge(df_pob_info, on=col_doc, how='left')
            cols_finales_f = [c for c in cols_ordenadas if c in df_realizadas_fuera_cohorte.columns]
            df_realizadas_fuera_cohorte = df_realizadas_fuera_cohorte[cols_finales_f].copy()

    # Muestra de conteos
    n_real_cohorte = len(df_realizadas_cohorte)
    n_real_fuera = len(df_realizadas_fuera_cohorte)
    n_pend_cohorte = len(df_pendientes_cohorte)

    print(f" -> Atenciones realizadas en COHORTE        : {n_real_cohorte:,}")
    print(f" -> Atenciones realizadas FUERA DE COHORTE  : {n_real_fuera:,} (Atenciones 'de más')")
    print(f" -> Atenciones pendientes en COHORTE        : {n_pend_cohorte:,}")

    return df_realizadas_cohorte, df_realizadas_fuera_cohorte, df_pendientes_cohorte

