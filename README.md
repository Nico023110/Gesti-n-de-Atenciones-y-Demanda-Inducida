# Plataforma Web RFAST & Gestión de Demanda Inducida

Sistema web interactivo y motor de auditoría de inconsistencias RIPS, atenciones de salud y demanda inducida para EPS y la ESE Ladera.

🌐 **Despliegue en Vercel**: [gesti-n-de-atenciones-y-demanda-ind.vercel.app](https://gesti-n-de-atenciones-y-demanda-ind.vercel.app/)

## 🚀 Características Principales

* **Interfaz Web en Vercel**: Carga de archivos Excel (`.xlsx`, `.csv`) directamente desde el navegador sin instalación previa.
* **Motor RFAST de 7 Validaciones**: Detección automática de errores en Causa Externa, Finalidad RIPS, Odontología, Planificación Familiar, Detección Temprana y Educación Individual.
* **Tableros e Indicadores KPI**: Métricas en tiempo real, gráficos de dona y barras por regla y por zona operativa.
* **Agrupación por Zonas Comunas y Rurales**: Clasificación automática en Comunas 01, 03, 17, 18, 20, Zona Rural Norte y Zona Rural Sur.
* **Exportación a Excel Formateado**: Generación de reportes Excel con celdas resaltadas en rojo y amarillo.
* **Soporte CLI / Consola Python**: Mantiene los scripts originales de consola en la carpeta `scripts/`.

## 📁 Estructura del Proyecto

```
Proyecto_Poblacion/
├── index.html                    # Aplicación Web principal (HTML5 + Lucide + Chart.js)
├── styles.css                    # Sistema de diseño moderno (Glassmorphism & Dark Mode)
├── app.js                        # Motor de auditoría y lógica cliente en JavaScript/SheetJS
├── vercel.json                   # Configuración para despliegue automatizado en Vercel
├── config/
│   └── config.py                 # Configuración Python
├── datos/
│   └── catalogos/                # Reglas y catálogos
└── scripts/                      # Scripts originales de consola Python
    ├── main.py
    ├── cruzar_poblacion.py
    └── motor_reglas.py
```

## ⚙️ Uso Local

### Servidor Web Local
Puedes abrir directamente el archivo `index.html` en tu navegador o servirlo con Python:
```bash
python -m http.server 8000
```
Luego navega a `http://localhost:8000`.

### Consola Python CLI
```bash
python scripts/main.py
```
