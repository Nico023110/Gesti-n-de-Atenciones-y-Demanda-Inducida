# -*- coding: utf-8 -*-
"""
Módulo: normalizar_columnas.py
Descripción: Normaliza los nombres de las columnas, traduce sinónimos/variaciones,
             unifica campos duplicados y preserva las columnas únicas de cada EAPB.
"""

import re
import unicodedata
import pandas as pd

# Diccionario de equivalencias para unificar variables con pequeñas variaciones
# Formato: 'nombre_estandar': ['variacion_1', 'variacion_2', ...]
# =============================================================================
# DICCIONARIO DE EQUIVALENCIAS Y SINÓNIMOS (MAPEO_ALIAS - 67 COLUMNAS MAESTRAS)
# =============================================================================

MAPEO_ALIAS = {
    # --- 1. IDENTIFICACIÓN Y DEMOGRAFÍA BÁSICA ---
    'primer_nombre': ['p_nombre', 'primer_nombre_1', 'nombre1', 'primernombre'],
    'segundo_nombre': ['s_nombre', 'segundo_nombre_1', 'nombre2', 'segundonombre'],
    'primer_apellido': ['p_apellido', 'primer_apellido_1', 'apellido1', 'primerapellido'],
    'segundo_apellido': ['s_apellido', 'segundo_apellido_1', 'apellido2', 'segundoapellido'],
    'tipo_identificacion': ['tipo_identificacion_1', 'tipo_doc', 'td', 'tipoidentificacion', 'tipodocumentoidentificacion', 'tipo_documento'],
    'nro_identificacion': ['nro_identificacion_1', 'num_documento', 'documento', 'id_paciente', 'numeroidentificacion', 'numdocumentoidentificacion'],
    
    # --- 2. AFILIACIÓN Y CURSO DE VIDA ---
    'fecha_nacimiento': ['fecha_nacimiento_1', 'fec_nacimiento', 'f_nac', 'fechanacimiento'],
    'edad_actual': ['edad_actual_1', 'edad', 'edad_anios'],
    'sexo': ['sexo_1', 'cod_sexo', 'genero', 'codsexo'],
    'regimen': ['regimen_1', 'tipo_regimen'],
    'etnia': ['poblacion_etnica', 'grupo_etnico', 'etnia_1'],
    'discapacidad': ['poblacion_discapacidad', 'discapacidad_1', 'tipo_discapacidad'],
    'victima_conflicto_armado': ['conflicto_armado', 'victima_conflicto_armado_1', 'poblacion_victima'],
    'curso_vida': ['curso_vida_1', 'ciclo_vida', 'momento_curso_vida'],
    'clasificacion_riesgo': ['clasificacion_riesgo_1', 'riesgo_cardiovascular', 'estratificacion'],
    
    # --- 3. ATENCIONES CLÍNICAS GENERALES Y ESPECIALIDADES ---
    'valoracion_integral': ['valoracion_integral_1', 'fecha_mas_reciente_de_vps_efectiva', 'fecha_de_vps_1_vez_mas_reciente', 'fecha_de_vps_mas_reciente', 'vps', 'vps_efectiva'],
    'atencion_medicina_general': ['atencion_medicina_general_1', 'medicina_general', 'val_medicina_general'],
    'atencion_enfermeria': ['atencion_enfermeria_1', 'enfermeria', 'consulta_enfermeria'],
    'atencion_pediatria': ['atencion_pediatria_1', 'pediatria', 'consulta_pediatria'],
    'atencion_odontologia': ['atencion_odontologia_1', 'odontologia', 'val_odontologica'],
    'aplicacion_fluor': ['aplicacion_fluor_1', 'fluor', 'topico_fluor'],
    'profilaxis': ['profilaxis_1', 'limpieza_dental', 'profilaxis_dental'],
    'control_placa': ['control_placa_1', 'placa_bacteriana', 'control_placa_dental'],
    'atencion_nutricion': ['atencion_nutricion_1', 'nutricion', 'val_nutricion'],
    'atencion_psicologia': ['atencion_psicologia_1', 'psicologia', 'val_psicologia'],
    'consulta_trabajo_social': ['consulta_trabajo_social_1', 'trabajo_social'],
    'medicina_interna': ['medicina_interna_1', 'consulta_medicina_interna'],
    
    # --- 4. PLANIFICACIÓN Y SALUD SEXUAL ---
    'planificacion_familiar': ['planificacon_familiar', 'planificacion_familiar_1', 'planificacon_familiar_1', 'c_planificacion'],
    'atencion_preconcepcional': ['atencion_preconcepcional_1', 'preconcepcional', 'val_preconcepcional'],
    'atencion_prenatal': ['atencion_prenatal_1', 'prenatal', 'control_prenatal'],
    'curso_maternidad': ['curso_maternidad_1', 'maternidad', 'psicoprofilactico'],
    'atencion_puerperio': ['atencion_puerperio_1', 'puerperio', 'control_puerperio'],
    
    # --- 5. TAMIZAJES DE CÁNCER ---
    'citologia': ['tamizaje_cancer_cuello_uterino', 'tamizaje_cancer_cuello_uterino_1', 'citologia_1', 'cito'],
    'colposcopia': ['colposcopia_1', 'examen_colposcopia'],
    'biopsia_cervical': ['biopsia_cervical_1', 'biopsia_cuello_uterino'],
    'mamografia': ['tamizaje_cancer_mama', 'tamizaje_cancer_mama_1', 'mamografia_1', 'mamo'],
    'biopsia_mama': ['biopsia_mama_1', 'examen_biopsia_mama'],
    'tamizaje_cancer_prostata': ['tamizaje_cancer_prostata_1', 'psa', 'tacto_rectal'],
    'biopsia_prostata': ['biopsia_prostata_1', 'examen_biopsia_prostata'],
    'tamizaje_cancer_colon': ['tamizaje_cancer_colon_1', 'sangre_oculta_heces', 'fobt'],
    'colonoscopia': ['colonoscopia_1', 'examen_colonoscopia'],
    'biopsia_colon': ['biopsia_colon_1', 'examen_biopsia_colon'],
    'asewgoria_genetica': ['asewgoria_genetica_1', 'asesoria_genetica'],
    
    # --- 6. TAMIZAJES CARDIOVASCULARES, METABÓLICOS Y OTROS ---
    'glicemia': ['tamizaje_glicemia', 'tamizaje_glicemia_1', 'glicemia_1', 'glucosa'],
    'hemoglobina_glicosilada': ['hemoglobina_glicosilada_1', 'hba1c'],
    'perfil_lipídico': ['tamizaje_perfil_lipídico', 'tamizaje_perfil_lipídico_1', 'perfil_lipidico', 'perfil_lipidico_1'],
    'creatinina': ['tamizaje_creatinina', 'tamizaje_creatinina_1', 'creatinina_1'],
    'albuminuria': ['tamizaje_albuminuria', 'tamizaje_albuminuria_1', 'albuminuria_1'],
    'prueba_vih': ['tamizaje_prueba_vih', 'tamizaje_prueba_vih_1', 'vih', 'prueba_vih_1'],
    'prueba_sifilis': ['tamizaje_prueba_sifilis', 'tamizaje_prueba_sifilis_1', 'sifilis', 'vdrl'],
    'prueba_hepatitis_b': ['tamizaje_prueba_hepatitis_b', 'tamizaje_prueba_hepatitis_b_1', 'hepatitis_b'],
    'prueba_hepatitis_c': ['tamizaje_prueba_hepatitis_c', 'tamizaje_prueba_hepatitis_c_1', 'hepatitis_c'],
    'hemoglobina_anemia': ['tamizaje_hemoglobina_anemia', 'tamizaje_hemoglobina_anemia_1', 'hemoglobina', 'anemia'],
    
    # --- 7. TAMIZAJES SENSORIALES Y DESARROLLO ---
    'agudeza_visual': ['tamizaje_agudeza_visual', 'tamizaje_agudeza_visual_1', 'agudezavisual'],
    'agudeza_auditiva': ['tamizaje_agudeza_auditiva', 'tamizaje_agudeza_auditiva_1', 'audiometria'],
    'evaluacion_desarrollo': ['tamizaje_evaluacion_desarrollo', 'tamizaje_evaluacion_desarrollo_1', 'escala_valleada'],
    'salud_mental': ['tamizaje_salud_mental', 'tamizaje_salud_mental_1', 'psicologia_tamizaje'],
    
    # --- 8. ESQUEMA DE VACUNACIÓN (BCG hasta VPH) ---
    'vacuna_bcg': ['vacuna_bcg_1', 'bcg'],
    'vacuna_hepatitis_b_rn': ['vacuna_hepatitis_b_rn_1', 'hep_b_recien_nacido'],
    'vacuna_polio': ['vacuna_polio_1', 'polio', 'vop'],
    'vacuna_pentavalente': ['vacuna_pentavalente_1', 'pentavalente'],
    'vacuna_rotavirus': ['vacuna_rotavirus_1', 'rotavirus'],
    'vacuna_neumococo': ['vacuna_neumococo_1', 'neumococo'],
    'vacuna_influenza': ['vacuna_influenza_1', 'influenza'],
    'vacuna_srp': ['vacuna_srp_1', 'srp', 'triple_viral'],
    'vacuna_fiebre_amarilla': ['vacuna_fiebre_amarilla_1', 'fiebre_amarilla'],
    'vacuna_hepatitis_a': ['vacuna_hepatitis_a_1', 'hepatitis_a'],
    'vacuna_varicela': ['vacuna_varicela_1', 'varicela'],
    'vacuna_dpt': ['vacuna_dpt_1', 'dpt'],
    'vacuna_vph': ['vacuna_vph_1', 'vph', 'papiloma'],
    'vacuna_toxoide_tetanico': ['vacuna_toxoide_tetanico_1', 'tt', 'td_vacuna'],
    'vacuna_covid19': ['vacuna_covid19_1', 'covid_19', 'vacuna_covid'],
    
    # --- NUEVO: UNIFICACIÓN DE CESIÓN Y FECHA DE CESIÓN ---
    'fecha_cesion': ['fecha_cesin', 'fecha_cesion', 'f_cesion', 'fec_cesion', 'fecha_cesion_1', 'f_cesin'],
    'cesion': ['cesin', 'cesion','cesion_1', 'cesin_1', 'estado_cesion', 'tipo_cesin'],

    # --- CONTACTO Y UBICACIÓN ---
    'direccion_residencia': ['direccionresidencia', 'direccion', 'direccion_residencia_1'],
    'barrio': ['barrio_residencia'],
    'comuna': ['localidadcomuna', 'localidad_comuna', 'comuna', 'localidad'],
    'celular': ['telefono_celular', 'num_celular'],
    'celular2': ['celular_2'],
    'telefono_fijo': ['telefonofijo', 'telefono', 'tel_fijo'],
    'correo_electronico': ['correoelectronico', 'email', 'correo'],
}


def normalizar_nombre(columna: str) -> str:
    """
    Convierte un nombre de columna a formato estándar en minúsculas y sin caracteres especiales.
    """
    columna = str(columna).strip()
    columna = unicodedata.normalize("NFKD", columna)
    columna = columna.encode("ascii", "ignore").decode("utf-8")
    columna = columna.lower()
    columna = re.sub(r"[^a-z0-9]+", "_", columna)
    columna = re.sub(r"_+", "_", columna).strip("_")

    # Traducir variaciones si la columna está en el diccionario de alias
    for nombre_estandar, lista_alias in MAPEO_ALIAS.items():
        if columna in lista_alias or columna == nombre_estandar:
            return nombre_estandar

    return columna


def normalizar_dataframe(df: pd.DataFrame):
    """
    1. Normaliza los nombres de todas las columnas y traduce alias/sinónimos.
    2. Si encuentra columnas con el mismo nombre normalizado, las unifica.
    """
    df_norm = df.copy()
    
    nombres_originales = list(df_norm.columns)
    df_norm.columns = [normalizar_nombre(col) for col in df_norm.columns]
    
    equivalencias = dict(zip(nombres_originales, df_norm.columns))

    # Unificación de columnas duplicadas dentro del mismo archivo
    if df_norm.columns.has_duplicates:
        df_norm = df_norm.T.groupby(level=0, sort=False).first().T
        print(" -> [UNIFICACIÓN] Se fusionaron columnas duplicadas dentro del archivo.")

    return df_norm, equivalencias