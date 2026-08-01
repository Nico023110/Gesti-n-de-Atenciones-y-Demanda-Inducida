# -*- coding: utf-8 -*-
"""
Módulo: actualizar_nominal.py
Descripción: Actualiza el DataFrame consolidado de la Nominal con las fechas
             de las actividades realizadas según el motor de reglas.
"""

import pandas as pd
from utilidades import separador


def crear_indice_pacientes(df_nominal: pd.DataFrame, columna_documento: str = "nro_identificacion") -> dict:
    """
    Crea un diccionario {documento: fila} para encontrar rápidamente
    el índice de un paciente dentro del DataFrame de la nominal.
    """
    if columna_documento not in df_nominal.columns:
        print(f" [ADVERTENCIA] No existe la columna '{columna_documento}' en la nominal.")
        return {}

    serie_doc = df_nominal[columna_documento].astype(str).str.strip().str.upper()
    indice = {doc: i for i, doc in enumerate(serie_doc) if doc and doc not in ['NAN', 'NONE', 'NAT', '']}
    return indice


def actualizar_nominal(df_nominal: pd.DataFrame, df_realizadas: pd.DataFrame) -> pd.DataFrame:
    """
    Actualiza la nominal con las fechas de las actividades realizadas.
    Para cada paciente con actividad realizada, escribe la fecha en la columna
    correspondiente de la nominal.
    """
    separador("ACTUALIZANDO BASE NOMINAL")

    if df_nominal is None or df_nominal.empty:
        print(" [ADVERTENCIA] La nominal está vacía o no fue cargada.")
        return df_nominal

    if df_realizadas is None or df_realizadas.empty:
        print(" [INFORMACIÓN] No hay actividades realizadas para actualizar.")
        return df_nominal

    df_res = df_nominal.copy()
    indice_pacientes = crear_indice_pacientes(df_res, "nro_identificacion")

    # Identificar columna(s) de fecha próxima atención / consulta en la nominal
    cols_proxima = [
        c for c in df_res.columns 
        if any(w in str(c).lower() for w in ['proxima', 'próxima', 'sig_atenc', 'siguiente_cons', 'prox_cons']) 
        and not any(ex in str(c).lower() for ex in ['anterior', 'mas_reciente', 'novedad', 'afiliacion', 'parto', 'nacimiento'])
    ]
    if not cols_proxima and 'fecha_de_proxima_consulta' in df_res.columns:
        cols_proxima = ['fecha_de_proxima_consulta']
    elif not cols_proxima:
        df_res['fecha_de_proxima_consulta'] = None
        cols_proxima = ['fecha_de_proxima_consulta']
        columnas_creadas = {'fecha_de_proxima_consulta'}
    else:
        columnas_creadas = set()

    actualizaciones_exitosas = 0
    actualizaciones_proxima = 0
    pacientes_proxima_actualizados = set()

    for _, fila in df_realizadas.iterrows():
        doc = str(fila['nro_identificacion']).strip().upper()
        col = str(fila['columna_nominal']).strip().lower()
        fecha = str(fila['fecha_atencion']).strip()

        try:
            freq_meses = int(float(str(fila.get('frecuencia_meses', 12)).strip()))
        except (ValueError, TypeError):
            freq_meses = 12

        if doc in indice_pacientes:
            idx = indice_pacientes[doc]

            # Crear la columna si no existe en la nominal
            if col not in df_res.columns:
                df_res[col] = None
                columnas_creadas.add(col)

            # Actualizar solo si la celda está vacía o la fecha nueva es más reciente
            val_actual = df_res.at[idx, col]
            hubo_actualizacion = False
            if pd.isna(val_actual) or str(val_actual).strip() in ['', 'nan', 'none', 'nat', 'None']:
                df_res.at[idx, col] = fecha
                actualizaciones_exitosas += 1
                hubo_actualizacion = True
            else:
                if str(fecha) > str(val_actual):
                    df_res.at[idx, col] = fecha
                    actualizaciones_exitosas += 1
                    hubo_actualizacion = True

            # Actualizar fecha de próxima atención / consulta en base a la frecuencia de la norma
            dt_atencion = pd.to_datetime(fecha, errors='coerce')
            if pd.notna(dt_atencion) and freq_meses > 0:
                dt_proxima = dt_atencion + pd.DateOffset(months=freq_meses)
                str_proxima = dt_proxima.strftime('%Y-%m-%d')

                for col_prox in cols_proxima:
                    val_prox_actual = df_res.at[idx, col_prox]
                    if idx not in pacientes_proxima_actualizados:
                        esta_vacio = pd.isna(val_prox_actual) or str(val_prox_actual).strip() in ['', 'nan', 'none', 'nat', 'None']
                        esta_vencido = not esta_vacio and str(val_prox_actual) < str(pd.Timestamp.now().strftime('%Y-%m-%d'))
                        if hubo_actualizacion or esta_vacio or esta_vencido or str_proxima > str(val_prox_actual):
                            df_res.at[idx, col_prox] = str_proxima
                            pacientes_proxima_actualizados.add(idx)
                            actualizaciones_proxima += 1
                    else:
                        if str_proxima < str(val_prox_actual):
                            df_res.at[idx, col_prox] = str_proxima

    print(f" -> Pacientes actualizados en la nominal : {actualizaciones_exitosas:,}")
    print(f" -> Fechas próxima atención actualizadas : {actualizaciones_proxima:,}")
    if columnas_creadas:
        print(f" -> Columnas de actividad creadas        : {len(columnas_creadas)} ({', '.join(sorted(columnas_creadas))})")

    return df_res