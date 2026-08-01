# Gestión de Atenciones y Demanda Inducida

Sistema de procesamiento, cruce y análisis de población de afiliados de salud contra bases nominales y de atenciones (FEV) para la gestión de demanda inducida por EPS (Coosalud, Emssanar, Nueva EPS, SOS, etc.).

## 📁 Estructura del Proyecto

```
Proyecto_Poblacion/
├── config/
│   └── config.py                 # Configuración de rutas estáticas y dinámicas por EPS y periodo
├── datos/
│   ├── catalogos/                # Reglas de actividades y catálogos de cruce
│   ├── fev/                      # Archivos FEV (Facturación / Atenciones)
│   └── Historial/                # Nominal acumulada e histórico de afiliados
├── logs/                         # Registros de ejecución
└── scripts/
    ├── main.py                   # Script principal de ejecución del pipeline
    ├── cargar_poblacion.py       # Carga y estructuración de base poblacional por EPS
    ├── cargar_fev.py             # Carga y procesamiento de atenciones FEV
    ├── cargar_nominal.py         # Carga de archivos nominales de afiliados
    ├── actualizar_nominal.py     # Actualización de bases nominales históricas
    ├── cruzar_poblacion.py       # Algoritmos de cruce poblacional vs atenciones
    ├── motor_reglas.py           # Motor de validación y reglas de negocio
    ├── formatear_excel.py        # Generación de reportes finales en Excel formateados
    ├── normalizar_columnas.py    # Limpieza y estandarización de nombres de campos
    └── utilidades.py             # Funciones auxiliares de soporte
```

## 🚀 Requisitos e Instalación

### Requisitos
* Python 3.8+

### Librerías Necesarias
```bash
pip install pandas openpyxl numpy
```

## ⚙️ Ejecución

Para iniciar el proceso interactivo de consolidación y demanda inducida:

```bash
python scripts/main.py
```

El script solicitará seleccionar el mes/periodo y la EPS correspondiente a procesar, ejecutando el flujo completo de validación, cruce y exportación de salida.
