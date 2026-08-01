# -*- coding: utf-8 -*-
"""
Módulo principal: main.py
Descripción: Orquesta el flujo del Motor de Actualización de Actividades.

Flujo Replanteado:
    1. Cargar Población asignada a la IPS
    2. Cargar FEV / RIPS (atenciones realizadas)
    3. Motor de Reglas en Población IPS × FEV → Actividades necesarias, realizadas y pendientes de la IPS
    4. Cargar Cohorte / Nominal (pacientes de cada EAPB)
    5. Cruzar & Clasificar → Atenciones en Cohorte vs. Atenciones Fuera de Cohorte ("atenciones de más")
    6. Actualizar Nominal EAPB con las atenciones de la cohorte
    7. Exportar resultados (Nominal actualizada, Pendientes cohorte, Realizadas fuera de cohorte, Resumen ejecutivo)
"""

import os
import pandas as pd

from utilidades import separador
from cargar_nominal import cargar_nominales
from cargar_poblacion import cargar_poblacion
from cruzar_poblacion import clasificar_atenciones_cohorte
from cargar_fev import cargar_fev
from motor_reglas import ejecutar_motor_reglas
from actualizar_nominal import actualizar_nominal
from formatear_excel import formatear_todos_los_excels, MAPA_COLUMNAS_PENDIENTES, MAPA_COLUMNAS_FUERA
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config')))
import config


def exportar_resultados(df_nominal_actualizada: pd.DataFrame,
                        df_realizadas_cohorte: pd.DataFrame,
                        df_realizadas_fuera_cohorte: pd.DataFrame,
                        df_pendientes_cohorte: pd.DataFrame,
                        df_necesarias_pob: pd.DataFrame):
    """
    Guarda los resultados del procesamiento en la carpeta datos/salida/.
    Genera reportes por EAPB, auditorías de cohorte, atenciones fuera de cohorte
    y resumen ejecutivo mejorado con múltiples hojas Excel.
    """
    separador("EXPORTANDO RESULTADOS A DATOS/SALIDA")

    if not os.path.exists(config.RUTA_SALIDA):
        os.makedirs(config.RUTA_SALIDA, exist_ok=True)

    # 1. Exportar auditoría de actividades realizadas en la Cohorte (Excel)
    if df_realizadas_cohorte is not None and not df_realizadas_cohorte.empty:
        ruta_cohorte_xlsx = os.path.join(config.RUTA_SALIDA, "actividades_realizadas_cohorte.xlsx")
        try:
            cols_drop = [c for c in ['columna_nominal', 'id_regla', 'frecuencia_meses'] if c in df_realizadas_cohorte.columns]
            df_export_cohorte = df_realizadas_cohorte.drop(columns=cols_drop).rename(columns=MAPA_COLUMNAS_FUERA)
            df_export_cohorte.to_excel(ruta_cohorte_xlsx, index=False)
            print(f" [OK] Actividades realizadas en Cohorte (Excel) guardadas en:\n      {ruta_cohorte_xlsx}")
        except Exception as e:
            print(f" [ADVERTENCIA] No se pudo guardar Excel de realizadas en cohorte: {e}")

    # 2. Exportar atenciones realizadas a pacientes FUERA DE COHORTE ("Atenciones de más")
    if df_realizadas_fuera_cohorte is not None and not df_realizadas_fuera_cohorte.empty:
        ruta_fuera_xlsx = os.path.join(config.RUTA_SALIDA, "actividades_realizadas_fuera_cohorte.xlsx")
        try:
            cols_drop = [c for c in ['columna_nominal', 'id_regla', 'frecuencia_meses'] if c in df_realizadas_fuera_cohorte.columns]
            df_export_fuera = df_realizadas_fuera_cohorte.drop(columns=cols_drop).rename(columns=MAPA_COLUMNAS_FUERA)
            df_export_fuera.to_excel(ruta_fuera_xlsx, index=False)
            print(f" [OK] Atenciones realizadas FUERA DE COHORTE (Excel) guardadas en:\n      {ruta_fuera_xlsx}")
        except Exception as e:
            print(f" [ADVERTENCIA] No se pudo guardar Excel de realizadas fuera de cohorte: {e}")

    # 3. Exportar actividades pendientes únicamente de la COHORTE
    if df_pendientes_cohorte is not None and not df_pendientes_cohorte.empty:
        ruta_xlsx = os.path.join(config.RUTA_SALIDA, "actividades_pendientes.xlsx")
        try:
            cols_drop_exp = [c for c in ['id_regla', 'columna_nominal', 'frecuencia_normativa'] if c in df_pendientes_cohorte.columns]
            df_export_pen = df_pendientes_cohorte.drop(columns=cols_drop_exp).rename(columns=MAPA_COLUMNAS_PENDIENTES)
            with pd.ExcelWriter(ruta_xlsx, engine='xlsxwriter') as writer:
                df_export_pen.to_excel(writer, index=False, sheet_name='Pendientes')
                workbook = writer.book
                worksheet = writer.sheets['Pendientes']
                
                for col_num, value in enumerate(df_export_pen.columns.values):
                    # Asignar color base
                    color = '#1F4E78' # NAVY
                    if value in ['Edad Paciente', 'Unidad Edad', 'Curso de Vida']: color = '#2E75B6'
                    elif value in ['Actividad Requerida', 'Rango Edad (Norma 3280)']: color = '#005B60'
                    elif value in ['Frecuencia según Edad', 'Detalle de la Atención Pendiente']: color = '#806000'
                    elif value in ['Celular Principal', 'Celular Secundario', 'Teléfono Fijo', 'Dirección Residencia', 'Barrio', 'Comuna', 'Correo Electrónico', 'IPS Asignada']: color = '#375623'
                    elif value == 'Estado de Gestión': color = '#595959'
                    
                    header_format = workbook.add_format({
                        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                        'fg_color': color, 'font_color': 'white', 'border': 1
                    })
                    
                    worksheet.write(0, col_num, value, header_format)
                    width = 48 if value == 'Detalle de la Atención Pendiente' else max(len(value) + 4, 14)
                    worksheet.set_column(col_num, col_num, width)
            
            print(f" [OK] Archivo de actividades pendientes de Cohorte (Excel con estilos nativos) guardado en:\n      {ruta_xlsx}")
        except Exception as e:
            print(f" [ADVERTENCIA] No se pudo guardar Excel de pendientes: {e}")

    # 4. Exportar Nominal Actualizada consolidada
    if df_nominal_actualizada is not None and not df_nominal_actualizada.empty:
        ruta_csv = os.path.join(config.RUTA_SALIDA, "nominal_consolidada_actualizada.csv")
        df_nominal_actualizada.to_csv(ruta_csv, index=False, sep=";", encoding="utf-8-sig")
        print(f" [OK] Nominal consolidada (CSV) guardada en:\n      {ruta_csv}")

        # 5. Exportar por EAPB individual en Excel
        if 'eapb' in df_nominal_actualizada.columns:
            for eapb, df_sub in df_nominal_actualizada.groupby('eapb'):
                nombre_limpio = "".join(c for c in eapb if c.isalnum() or c in (' ', '_')).strip()
                ruta_eapb = os.path.join(config.RUTA_SALIDA, f"Nominal_Actualizada_{nombre_limpio}.xlsx")
                try:
                    df_sub.to_excel(ruta_eapb, index=False)
                    print(f" [OK] Nominal para {eapb} ({len(df_sub):,} pacientes) guardada en:\n      {ruta_eapb}")
                except Exception as e:
                    print(f" [ADVERTENCIA] No se pudo guardar Excel para {eapb}: {e}")

    # =========================================================================
    # 6. RESUMEN EJECUTIVO MEJORADO (Excel con múltiples hojas)
    # =========================================================================
    _generar_resumen_ejecutivo(
        df_necesarias_pob, df_realizadas_cohorte,
        df_realizadas_fuera_cohorte, df_pendientes_cohorte
    )


def _generar_resumen_ejecutivo(df_necesarias_pob, df_realizadas_cohorte,
                                df_realizadas_fuera_cohorte, df_pendientes_cohorte):
    """
    Genera resumen_gestion.xlsx con 3 hojas:
      1. Resumen por Actividad
      2. Resumen por Curso de Vida
      3. Indicadores Globales
    """
    ruta_resumen = os.path.join(config.RUTA_SALIDA, "resumen_gestion.xlsx")

    # =====================================================================
    # HOJA 1: RESUMEN POR ACTIVIDAD
    # =====================================================================
    dict_reglas = {}
    # Recopilar info de reglas con frecuencia
    if df_necesarias_pob is not None and not df_necesarias_pob.empty:
        for _, row in df_necesarias_pob[['id_regla', 'actividad', 'frecuencia_meses']].drop_duplicates(subset=['id_regla']).iterrows():
            dict_reglas[row['id_regla']] = {
                'actividad': row['actividad'],
                'frecuencia_meses': row['frecuencia_meses']
            }

    resumen_actividad = []
    for id_regla in sorted(dict_reglas.keys()):
        info = dict_reglas[id_regla]
        actividad = info['actividad']
        freq = info['frecuencia_meses']

        # Etiqueta legible de frecuencia
        if freq <= 1:
            freq_label = "Mensual"
        elif freq <= 3:
            freq_label = "Trimestral"
        elif freq <= 6:
            freq_label = "Semestral"
        elif freq <= 12:
            freq_label = "Anual"
        elif freq <= 24:
            freq_label = "Cada 2 años"
        elif freq <= 36:
            freq_label = "Cada 3 años"
        elif freq <= 60:
            freq_label = "Cada 5 años"
        else:
            freq_label = f"Cada {freq} meses"

        cnt_ips = len(df_necesarias_pob[df_necesarias_pob['id_regla'] == id_regla]) if df_necesarias_pob is not None and not df_necesarias_pob.empty else 0
        cnt_rc = len(df_realizadas_cohorte[df_realizadas_cohorte['id_regla'] == id_regla]) if df_realizadas_cohorte is not None and not df_realizadas_cohorte.empty else 0
        cnt_pc = len(df_pendientes_cohorte[df_pendientes_cohorte['id_regla'] == id_regla]) if df_pendientes_cohorte is not None and not df_pendientes_cohorte.empty else 0
        cnt_esp = cnt_rc + cnt_pc
        cnt_rf = len(df_realizadas_fuera_cohorte[df_realizadas_fuera_cohorte['id_regla'] == id_regla]) if df_realizadas_fuera_cohorte is not None and not df_realizadas_fuera_cohorte.empty else 0
        total_r = cnt_rc + cnt_rf

        pct_c = round((cnt_rc / cnt_esp * 100), 1) if cnt_esp > 0 else 0.0
        pct_f = round((cnt_rf / total_r * 100), 1) if total_r > 0 else 0.0

        # Semáforo de cumplimiento
        if pct_c >= 80:
            semaforo = "🟢 ALTO"
        elif pct_c >= 50:
            semaforo = "🟡 MEDIO"
        elif pct_c >= 20:
            semaforo = "🟠 BAJO"
        else:
            semaforo = "🔴 CRITICO"

        resumen_actividad.append({
            'ID Regla': id_regla,
            'Actividad': actividad,
            'Frecuencia Norma (meses)': freq,
            'Periodicidad': freq_label,
            'Pob. IPS Necesita': cnt_ips,
            'Cohorte Esperados': cnt_esp,
            'Cohorte Realizados (Vigentes)': cnt_rc,
            'Cohorte Pendientes': cnt_pc,
            'Fuera Cohorte Realizados': cnt_rf,
            'Total Realizados IPS': total_r,
            '% Cumplimiento Cohorte': pct_c,
            '% Aporte Fuera Cohorte': pct_f,
            'Semaforo': semaforo
        })

    df_h1 = pd.DataFrame(resumen_actividad) if resumen_actividad else pd.DataFrame()

    # =====================================================================
    # HOJA 2: RESUMEN POR CURSO DE VIDA
    # =====================================================================
    resumen_cv = []
    cursos = ['RECIÉN NACIDO', 'PRIMERA INFANCIA', 'INFANCIA', 'ADOLESCENCIA', 'JUVENTUD', 'ADULTEZ', 'VEJEZ']

    for cv in cursos:
        cnt_nec = len(df_necesarias_pob[df_necesarias_pob['curso_vida'] == cv]) if df_necesarias_pob is not None and not df_necesarias_pob.empty and 'curso_vida' in df_necesarias_pob.columns else 0

        cnt_pc = 0
        if df_pendientes_cohorte is not None and not df_pendientes_cohorte.empty and 'curso_vida' in df_pendientes_cohorte.columns:
            cnt_pc = len(df_pendientes_cohorte[df_pendientes_cohorte['curso_vida'] == cv])

        # Pacientes únicos que necesitan en este curso de vida
        pac_unicos = 0
        if df_necesarias_pob is not None and not df_necesarias_pob.empty and 'curso_vida' in df_necesarias_pob.columns:
            pac_unicos = df_necesarias_pob[df_necesarias_pob['curso_vida'] == cv]['nro_identificacion'].nunique()

        # Actividades distintas que aplican
        n_actividades = 0
        if df_necesarias_pob is not None and not df_necesarias_pob.empty and 'curso_vida' in df_necesarias_pob.columns:
            n_actividades = df_necesarias_pob[df_necesarias_pob['curso_vida'] == cv]['id_regla'].nunique()

        resumen_cv.append({
            'Curso de Vida': cv,
            'Pacientes Unicos': pac_unicos,
            'Actividades Aplicables': n_actividades,
            'Total Atenc. Necesarias': cnt_nec,
            'Pendientes Cohorte': cnt_pc,
            'Brecha': cnt_pc
        })

    df_h2 = pd.DataFrame(resumen_cv) if resumen_cv else pd.DataFrame()

    # =====================================================================
    # HOJA 3: INDICADORES GLOBALES
    # =====================================================================
    total_necesarias = len(df_necesarias_pob) if df_necesarias_pob is not None else 0
    total_rc = len(df_realizadas_cohorte) if df_realizadas_cohorte is not None else 0
    total_rf = len(df_realizadas_fuera_cohorte) if df_realizadas_fuera_cohorte is not None else 0
    total_pc = len(df_pendientes_cohorte) if df_pendientes_cohorte is not None else 0
    total_realizadas = total_rc + total_rf
    pct_global = round((total_rc / (total_rc + total_pc) * 100), 1) if (total_rc + total_pc) > 0 else 0.0
    pct_fuera = round((total_rf / total_realizadas * 100), 1) if total_realizadas > 0 else 0.0

    # Top 5 actividades con mayor brecha
    top_brecha = []
    if df_h1 is not None and not df_h1.empty:
        top = df_h1.nlargest(5, 'Cohorte Pendientes')
        for _, r in top.iterrows():
            top_brecha.append({
                'Indicador': f"Brecha: {r['Actividad']}",
                'Valor': f"{int(r['Cohorte Pendientes']):,} pendientes"
            })

    indicadores = [
        {'Indicador': 'Fecha de Procesamiento', 'Valor': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')},
        {'Indicador': '-' * 40, 'Valor': '-' * 20},
        {'Indicador': 'Total Atenciones Necesarias (Población IPS)', 'Valor': f"{total_necesarias:,}"},
        {'Indicador': 'Atenciones Realizadas en Cohorte (Vigentes)', 'Valor': f"{total_rc:,}"},
        {'Indicador': 'Atenciones Realizadas Fuera de Cohorte', 'Valor': f"{total_rf:,}"},
        {'Indicador': 'Total Atenciones Realizadas IPS', 'Valor': f"{total_realizadas:,}"},
        {'Indicador': 'Atenciones Pendientes en Cohorte', 'Valor': f"{total_pc:,}"},
        {'Indicador': '-' * 40, 'Valor': '-' * 20},
        {'Indicador': '% Cumplimiento Global Cohorte', 'Valor': f"{pct_global}%"},
        {'Indicador': '% Atenciones Fuera de Cohorte', 'Valor': f"{pct_fuera}%"},
        {'Indicador': '-' * 40, 'Valor': '-' * 20},
        {'Indicador': 'TOP 5 ACTIVIDADES CON MAYOR BRECHA', 'Valor': ''},
    ] + top_brecha

    df_h3 = pd.DataFrame(indicadores)

    # =====================================================================
    # ESCRIBIR EXCEL CON MÚLTIPLES HOJAS
    # =====================================================================
    try:
        with pd.ExcelWriter(ruta_resumen, engine='openpyxl') as writer:
            if not df_h1.empty:
                df_h1.to_excel(writer, sheet_name='Resumen por Actividad', index=False)
            if not df_h2.empty:
                df_h2.to_excel(writer, sheet_name='Resumen por Curso de Vida', index=False)
            df_h3.to_excel(writer, sheet_name='Indicadores Globales', index=False)

        print(f" [OK] Resumen ejecutivo (Excel con 3 hojas) guardado en:\n      {ruta_resumen}")
    except Exception as e:
        print(f" [ADVERTENCIA] No se pudo guardar resumen Excel: {e}")
        # Fallback a CSV
        ruta_csv = os.path.join(config.RUTA_SALIDA, "resumen_gestion.csv")
        if not df_h1.empty:
            df_h1.to_csv(ruta_csv, index=False, sep=";", encoding="utf-8-sig")
            print(f" [OK] Resumen (fallback CSV) guardado en:\n      {ruta_csv}")

    # =========================================================================
    # FORMATEO VISUAL INTEGRAL Y ENCABEZADOS CLAROS (MODULO SEPARADO)
    # =========================================================================
    cols_actividad = (
        df_realizadas_cohorte['columna_nominal'].dropna().unique().tolist()
        if df_realizadas_cohorte is not None and not df_realizadas_cohorte.empty and 'columna_nominal' in df_realizadas_cohorte.columns
        else []
    )
    formatear_todos_los_excels(config.RUTA_SALIDA, cols_actividad)



def main():
    import config
    import sys
    separador("MOTOR DE ACTUALIZACIÓN Y AUDITORÍA DE ACTIVIDADES DE SALUD")

    base_datos = os.path.join(config.BASE_DIR, "datos")
    
    import re
    # Permitir especificar un periodo y EPS como argumentos (ej. python main.py 2026_06 EMSSANAR)
    periodo_filtro = sys.argv[1] if len(sys.argv) > 1 else None
    eps_filtro = sys.argv[2] if len(sys.argv) > 2 else None

    # Buscar automáticamente todas las carpetas de período/mes (ej. 2026_07, 2026_08, etc.)
    carpetas_meses = [
        d for d in os.listdir(base_datos)
        if os.path.isdir(os.path.join(base_datos, d)) and re.match(r'^\d{4}_\d{2}$', d)
    ]
    if periodo_filtro:
        carpetas_meses = [d for d in carpetas_meses if d == periodo_filtro]
    
    if not carpetas_meses:
        print(f" [ERROR] No se encontraron carpetas de períodos en {base_datos}")
        return

    # Procesar cada período encontrado
    for mes_corte in sorted(carpetas_meses):
        ruta_mes = os.path.join(base_datos, mes_corte)
        carpetas_eps = [
            d for d in os.listdir(ruta_mes)
            if os.path.isdir(os.path.join(ruta_mes, d)) and d not in ['fev_global', 'catalogos']
        ]
        if eps_filtro:
            carpetas_eps = [d for d in carpetas_eps if d.upper() == eps_filtro.upper()]
        
        if not carpetas_eps:
            print(f" [ADVERTENCIA] No se encontraron carpetas de EPS en {ruta_mes}")
            continue
            
        for eps in carpetas_eps:
            separador(f"PROCESANDO EPS: {eps} - PERÍODO: {mes_corte}")
            # set_rutas_dinamicas crea automáticamente las carpetas poblacion, nominal, fev, salida si no existen
            config.set_rutas_dinamicas(ruta_mes, eps)

            # =========================================================================
            # PASO 1: Cargar Población asignada a la IPS
            # =========================================================================
            df_poblacion = cargar_poblacion()
            if df_poblacion is None:
                print(f" [ADVERTENCIA] No se pudo cargar la población para {eps}. Saltando EPS.")
                continue
        
            # =========================================================================
            # PASO 2: Cargar FEV / RIPS (atenciones realizadas)
            # =========================================================================
            df_fev = cargar_fev()
        
            # =========================================================================
            # PASO 3: Cargar Cohorte / Nominal de las EAPB
            # =========================================================================
            df_nominal = cargar_nominales()
            if df_nominal is None:
                print(" [ADVERTENCIA] No se cargaron nominales. No habrá cruce de cohorte.")
                df_nominal = pd.DataFrame()
                docs_cohorte = set()
            else:
                docs_cohorte = set(df_nominal['nro_identificacion'].astype(str).str.strip().str.upper()) if 'nro_identificacion' in df_nominal.columns else set()
        
            # =========================================================================
            # PASO 4: Motor de Reglas → Evaluar Población IPS × FEV
            # =========================================================================
            df_necesarias_pob, df_realizadas_pob, df_pendientes_pob = ejecutar_motor_reglas(df_poblacion, df_fev, docs_cohorte, df_nominal)
        
            # =========================================================================
            # PASO 5: Cruzar & Clasificar (Atenciones en Cohorte vs. Fuera de Cohorte)
            # =========================================================================
            df_realizadas_cohorte, df_realizadas_fuera_cohorte, df_pendientes_cohorte = clasificar_atenciones_cohorte(
                df_necesarias_pob, df_realizadas_pob, df_pendientes_pob, df_nominal, df_poblacion
            )
        
            # =========================================================================
            # PASO 6: Actualizar Nominal EAPB con las atenciones de la cohorte
            # =========================================================================
            df_nominal_actualizada = actualizar_nominal(df_nominal, df_realizadas_cohorte)
        
            # =========================================================================
            # PASO 7: Exportar Resultados
            # =========================================================================
            exportar_resultados(
                df_nominal_actualizada,
                df_realizadas_cohorte,
                df_realizadas_fuera_cohorte,
                df_pendientes_cohorte,
                df_necesarias_pob
            )
        
            separador(f"PROCESO FINALIZADO PARA {eps} CON ÉXITO")

    separador("PROCESO GLOBAL FINALIZADO CON ÉXITO")


if __name__ == "__main__":
    main()