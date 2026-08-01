# =============================================================================
# CARGAR FEV
# =============================================================================

import os
import pandas as pd

import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config')))
import config
from utilidades import separador
from normalizar_columnas import normalizar_dataframe


def cargar_fev():

    separador("Leyendo archivos FEV")

    archivos = [
        archivo
        for archivo in os.listdir(config.RUTA_FEV)
        if archivo.lower().endswith(".csv")
    ]

    if len(archivos) == 0:
        print("No se encontraron archivos FEV.")
        return None

    lista_fev = []

    for archivo in archivos:

        ruta = os.path.join(config.RUTA_FEV, archivo)

        print(f"Leyendo: {archivo}")

        try:

            # Leer todo como texto
            df = pd.read_csv(
                ruta,
                sep=";",
                dtype=str,
                encoding="utf-8",
                low_memory=False
            )

        except UnicodeDecodeError:

            df = pd.read_csv(
                ruta,
                sep=";",
                dtype=str,
                encoding="latin-1",
                low_memory=False
            )

        # Normalizar nombres de columnas
        df, equivalencias = normalizar_dataframe(df)

        # Agregar columnas técnicas
        df["archivo_fev"] = os.path.splitext(archivo)[0]
        df["empresa"] = ""
        df["actividad_identificada"] = ""
        df["regla_aplicada"] = ""
        df["actualizada"] = False

        lista_fev.append(df)

        print(f"   Registros: {len(df):,}")

    df_fev = pd.concat(lista_fev, ignore_index=True)

    print("-" * 80)
    print(f"Total registros : {len(df_fev):,}")
    print(f"Total columnas  : {len(df_fev.columns)}")

    return df_fev