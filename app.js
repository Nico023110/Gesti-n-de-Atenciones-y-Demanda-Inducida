// Initialize Lucide icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// App State
const state = {
    records: [],
    inconsistencies: [],
    cleanedCount: 0,
    zonasStats: {},
    reglasStats: {},
    chartReglas: null,
    chartZonas: null
};

// Zone Definitions
const ZONAS_MAP = {
    'ZonaComuna-01': ['TERRON COLORADO', 'BELLAVISTA', 'VISTAHERMOSA', 'VISTA HERMOSA', 'LA PAZ'],
    'ZonaComuna-03': ['FRAY DAMIAN', 'CAÑA VERALEJO', 'CANAVERALEJO'],
    'ZonaComuna-17': ['PRIMERO DE MAYO'],
    'ZonaComuna-18': ['MELENDEZ', 'ALTO POLVORINES', 'ALTO NAPOLES', 'NAPOLES', 'POLVORINES', 'LOURDES'],
    'ZonaComuna-20': ['SILOE', 'BELEN', 'BRISAS DE MAYO', 'LA ESTRELLA', 'LA SIRENA', 'LA SULTANA'],
    'ZonaRuralNorte': ['MONTEBELLO', 'EL SALADITO', 'LA ELVIRA', 'FELIDIA', 'PENAS BLANCAS', 'PICHINDE', 'GOLONDRINAS', 'LA LEONERA', 'LA CASTILLA', 'LOS ANDES'],
    'ZonaRuralSur': ['LA BUITRERA', 'VILLACARMELO', 'PANCE', 'LA VORAGINE', 'EL HORMIGUERO', 'CASCAJAL']
};

function getZonaPorIPS(nombreIPS) {
    if (!nombreIPS) return 'Otras Sedes';
    const ipsUpper = nombreIPS.toUpperCase();
    for (const [zona, ipsList] of Object.entries(ZONAS_MAP)) {
        if (ipsList.some(ips => ipsUpper.includes(ips))) {
            return zona;
        }
    }
    return 'Zona General';
}

// Navigation Tabs
function initApp() {
    const navButtons = document.querySelectorAll('.nav-btn');
    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetTab = btn.getAttribute('data-tab');
            
            navButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.add('hidden');
            });
            document.getElementById(`tab-${targetTab}`).classList.remove('hidden');

            const pageTitles = {
                'dashboard': ['Panel de Auditoría & Gestión de Atenciones', 'Carga tus archivos Excel de RIPS o Población para ejecutar el motor de análisis en tiempo real.'],
                'inconsistencias': ['Motor RFAST - Inconsistencias RIPS', 'Detalle técnico de las 7 validaciones automáticas de consistencia clínica.'],
                'demanda': ['Gestión de Demanda Inducida', 'Cruce de afiliados nominales vs facturación de atenciones por EPS.'],
                'reglas': ['Zonas Operativas & Catálogos', 'Estructura geográfica de distribución por comunas y zona rural.']
            };
            if (pageTitles[targetTab]) {
                document.getElementById('page-title').textContent = pageTitles[targetTab][0];
                document.getElementById('page-subtitle').textContent = pageTitles[targetTab][1];
            }
        });
    });

    // File Dropzone Listeners
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('file-input');

    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // Demo Data Buttons
    document.getElementById('btn-load-demo-rfast').addEventListener('click', loadDemoRFAST);
    document.getElementById('btn-load-demo-poblacion').addEventListener('click', loadDemoPoblacion);

    // Search Input Listener
    document.getElementById('search-input').addEventListener('input', (e) => {
        filterTable(e.target.value);
    });

    // Export Button Listener
    document.getElementById('btn-export-excel').addEventListener('click', exportExcelReport);
}

// Handle Uploaded File
function handleFileSelect(file) {
    showLoader(`Procesando archivo: ${file.name}...`);
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
            
            processData(jsonData);
        } catch (err) {
            alert('Error al leer el archivo Excel/CSV. Verifica que sea un formato válido.');
            hideLoader();
        }
    };
    reader.readAsArrayBuffer(file);
}

// Main RFAST & Demanda Rule Engine
function processData(rows) {
    state.records = rows;
    state.inconsistencies = [];
    state.zonasStats = {};
    state.reglasStats = {
        'Regla 01 (Causa PyM)': 0,
        'Regla 02 (Control)': 0,
        'Regla 03 (Primera Vez)': 0,
        'Regla 04 (Odontología)': 0,
        'Regla 05 (Planificación)': 0,
        'Regla 06 (Detección Temprana)': 0,
        'Regla 07 (Educación Ind.)': 0
    };

    rows.forEach((r, idx) => {
        // Standardize Column Keys
        const getVal = (keys) => {
            for (let k of keys) {
                for (let key in r) {
                    if (key.trim().toLowerCase() === k.toLowerCase()) return String(r[key]).trim();
                }
            }
            return '';
        };

        const codigo = getVal(['codigo', 'cod_actividad']);
        const nombre = getVal(['nombre', 'actividad', 'nombre_actividad']);
        const finalidad = getVal(['finalidad', 'finalidad_rips', 'finalidad_']);
        const cexterna = getVal(['cexterna', 'causa_externa', 'causaexterna']);
        const centroprod = getVal(['centroprod', 'centro_produccion']);
        const nombreCentro = getVal(['nombre_centroproduccion', 'nombre_cen']);
        const prestador = getVal(['nombre_prestador', 'nombre_pre', 'ips', 'sede']);
        const documento = getVal(['documento', 'num_documento', 'cedula']);
        const factura = getVal(['factura', 'num_factura', 'fev']);

        const nombreUpper = nombre.toUpperCase();
        const centroUpper = nombreCentro.toUpperCase();

        let errorMsg = null;
        let reglaNombre = null;

        // Rule 01: Causa Externa Incorrecta PyM
        if (cexterna === '38' && finalidad === '11' && (centroUpper.includes('CURSO DE VIDA') || centroUpper.includes('PLANIFICACION FAMILIAR')) && (nombreUpper.includes('PRIMERA VEZ') || nombreUpper.includes('SEGUIMIENTO'))) {
            errorMsg = "La causa externa no puede ser 38 (Enfermedad General), debe ser 40 (Promoción y mantenimiento).";
            reglaNombre = 'Regla 01 (Causa PyM)';
        }
        // Rule 02: Finalidad en Programas de Control
        else if (['1415', '1416', '1417'].some(c => centroprod.includes(c)) && (nombreUpper.includes('CONTROL') || nombreUpper.includes('SEGUIMIENTO')) && !['16', '17', '23', '0', '28'].includes(finalidad)) {
            errorMsg = "La finalidad debe ser 28 (Tratamiento) por seguimiento de pacientes con diagnósticos definidos.";
            reglaNombre = 'Regla 02 (Control)';
        }
        // Rule 03: Finalidad en Primera Vez
        else if (nombreUpper.includes('PRIMERA VEZ') && ['1415', '1416', '1417'].includes(centroprod) && !['15', '23', '0'].includes(finalidad)) {
            errorMsg = "La finalidad debe ser 27 (Diagnóstico) al ser consulta por primera vez.";
            reglaNombre = 'Regla 03 (Primera Vez)';
        }
        // Rule 04: Odontología Control
        else if (nombreUpper.includes('CONTROL') && ['1300', '1303'].includes(centroprod) && !['16', '17', '0', '23'].includes(finalidad)) {
            errorMsg = "Debe ser finalidad 28 (Tratamiento) para continuidad clínica en odontología.";
            reglaNombre = 'Regla 04 (Odontología)';
        }
        // Rule 05: Planificación Familiar
        else if (nombreUpper.includes('CONSULTA') && centroprod === '1405' && !['19', '21', '23', '25', '0'].includes(finalidad)) {
            errorMsg = "Debe ser finalidad 31 (Planificación Familiar) y causa externa 40 (PYP PF).";
            reglaNombre = 'Regla 05 (Planificación)';
        }
        // Rule 06: Detección Temprana
        else if (nombreUpper.includes('CONSULTA') && ['1408', '1409', '1440', '1439'].includes(centroprod) && !['12', '15', '16', '23', '0'].includes(finalidad)) {
            errorMsg = "La finalidad debe ser 24 (Detección temprana de enfermedad general).";
            reglaNombre = 'Regla 06 (Detección Temprana)';
        }
        // Rule 07: Educación Individual
        else if (nombreUpper.includes('EDUCACION INDIVIDUAL') && !['0', '19', '20', '23', '28', '29', '30', '32', '33', '34', '38', '39', '40', '41', '42'].includes(finalidad)) {
            errorMsg = "Debe registrarse con finalidad de Promoción de la Salud (40 a 54).";
            reglaNombre = 'Regla 07 (Educación Ind.)';
        }

        const zona = getZonaPorIPS(prestador);

        if (errorMsg) {
            state.inconsistencies.push({
                codigo,
                nombre,
                finalidad,
                cexterna,
                prestador: prestador || 'SEDE PRINCIPAL',
                documento: documento || 'SIN DOC',
                factura: factura || `FACT-${idx + 1}`,
                inconsistencia: errorMsg,
                zona,
                regla: reglaNombre
            });

            state.reglasStats[reglaNombre] = (state.reglasStats[reglaNombre] || 0) + 1;
            state.zonasStats[zona] = (state.zonasStats[zona] || 0) + 1;
        }
    });

    state.cleanedCount = rows.length - state.inconsistencies.length;
    updateDashboardUI();
    hideLoader();
}

// Update UI Components
function updateDashboardUI() {
    const total = state.records.length;
    const errCount = state.inconsistencies.length;
    const cleanCount = state.cleanedCount;

    document.getElementById('kpi-total').textContent = total.toLocaleString();
    document.getElementById('kpi-total-sub').textContent = `${total} registros leídos`;

    document.getElementById('kpi-inconsistencias').textContent = errCount.toLocaleString();
    document.getElementById('kpi-inconsistencias-sub').textContent = total ? `${((errCount / total) * 100).toFixed(1)}% del total` : '0%';

    document.getElementById('kpi-conformes').textContent = cleanCount.toLocaleString();
    document.getElementById('kpi-conformes-sub').textContent = total ? `${((cleanCount / total) * 100).toFixed(1)}% sin error` : '0%';

    const numZonas = Object.keys(state.zonasStats).length;
    document.getElementById('kpi-zonas').textContent = numZonas || '7';

    document.getElementById('badge-count').textContent = `${errCount} errores`;
    document.getElementById('btn-export-excel').disabled = errCount === 0;

    renderTable(state.inconsistencies);
    renderCharts();
}

// Render Results Table
function renderTable(data) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (!data.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="empty-table-msg">
                    <i data-lucide="check-circle-2" style="color: var(--success)"></i>
                    <p>No se encontraron inconsistencias en la muestra actual.</p>
                </td>
            </tr>
        `;
        lucide.createIcons();
        return;
    }

    data.slice(0, 100).forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><code>${row.codigo || 'N/A'}</code></td>
            <td><strong>${row.nombre || 'Consulta General'}</strong></td>
            <td><span class="badge" style="background: rgba(99,102,241,0.15); color: #818CF8">${row.finalidad || '0'}</span></td>
            <td><span class="badge" style="background: rgba(245,158,11,0.15); color: #FBBF24">${row.cexterna || '38'}</span></td>
            <td>${row.prestador}</td>
            <td>${row.documento}</td>
            <td>${row.factura}</td>
            <td class="text-error">${row.inconsistencia}</td>
        `;
        tbody.appendChild(tr);
    });
}

function filterTable(query) {
    const q = query.toLowerCase();
    const filtered = state.inconsistencies.filter(row => {
        return row.codigo.toLowerCase().includes(q) ||
               row.nombre.toLowerCase().includes(q) ||
               row.documento.toLowerCase().includes(q) ||
               row.prestador.toLowerCase().includes(q) ||
               row.factura.toLowerCase().includes(q) ||
               row.inconsistencia.toLowerCase().includes(q);
    });
    renderTable(filtered);
}

// Render Visual Charts
function renderCharts() {
    // Chart 1: Reglas
    const ctx1 = document.getElementById('chart-reglas').getContext('2d');
    if (state.chartReglas) state.chartReglas.destroy();

    state.chartReglas = new Chart(ctx1, {
        type: 'doughnut',
        data: {
            labels: Object.keys(state.reglasStats),
            datasets: [{
                data: Object.values(state.reglasStats),
                backgroundColor: ['#6366F1', '#EF4444', '#10B981', '#F59E0B', '#3B82F6', '#A855F7', '#EC4899']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { color: '#94A3B8' } }
            }
        }
    });

    // Chart 2: Zonas
    const ctx2 = document.getElementById('chart-zonas').getContext('2d');
    if (state.chartZonas) state.chartZonas.destroy();

    state.chartZonas = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: Object.keys(state.zonasStats),
            datasets: [{
                label: 'Errores por Zona',
                data: Object.values(state.zonasStats),
                backgroundColor: 'rgba(99, 102, 241, 0.7)',
                borderColor: '#6366F1',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94A3B8' } },
                y: { ticks: { color: '#94A3B8' } }
            },
            plugins: {
                legend: { display: false }
            }
        }
    });
}

// Load Demo Dataset RFAST
function loadDemoRFAST() {
    showLoader('Cargando registros demostrativos RFAST...');
    setTimeout(() => {
        const demoData = [
            { codigo: '890201', nombre: 'CONSULTA DE PRIMERA VEZ POR MEDICINA GENERAL', finalidad: '11', cexterna: '38', centroprod: '1415', nombre_centroproduccion: 'CURSO DE VIDA', nombre_prestador: 'C.S. TERRON COLORADO', documento: '1144123456', factura: 'FEV-1001' },
            { codigo: '890301', nombre: 'CONSULTA DE CONTROL O SEGUIMIENTO', finalidad: '10', cexterna: '38', centroprod: '1415', nombre_centroproduccion: 'CONTROL HIPERTENSION', nombre_prestador: 'HOSPITAL CAÑA VERALEJO', documento: '1144654321', factura: 'FEV-1002' },
            { codigo: '890201', nombre: 'CONSULTA PRIMERA VEZ ODONTOLOGIA', finalidad: '11', cexterna: '38', centroprod: '1300', nombre_centroproduccion: 'ODONTOLOGIA', nombre_prestador: 'C.S. SILOE', documento: '31987654', factura: 'FEV-1003' },
            { codigo: '890202', nombre: 'CONSULTA DE PLANIFICACION FAMILIAR PRIMERA VEZ', finalidad: '10', cexterna: '38', centroprod: '1405', nombre_centroproduccion: 'PLANIFICACION FAMILIAR', nombre_prestador: 'C.S. MELENDEZ', documento: '1144998877', factura: 'FEV-1004' },
            { codigo: '890203', nombre: 'CONSULTA DETECCION TEMPRANA CANCER DE CUTERINO', finalidad: '11', cexterna: '38', centroprod: '1408', nombre_centroproduccion: 'DETECCION TEMPRANA', nombre_prestador: 'P.S. VISTAHERMOSA', documento: '66998877', factura: 'FEV-1005' },
            { codigo: '890205', nombre: 'CONSULTA DE CONTROL POR ODONTOLOGIA GENERAL', finalidad: '15', cexterna: '38', centroprod: '1303', nombre_centroproduccion: 'ODONTOLOGIA', nombre_prestador: 'P.S. MONTEBELLO', documento: '1144112233', factura: 'FEV-1006' },
            { codigo: '890201', nombre: 'EDUCACION INDIVIDUAL EN SALUD', finalidad: '10', cexterna: '38', centroprod: '1400', nombre_centroproduccion: 'PROMOCION', nombre_prestador: 'P.S. LA BUITRERA', documento: '94554433', factura: 'FEV-1007' }
        ];
        processData(demoData);
    }, 600);
}

// Load Demo Dataset Población
function loadDemoPoblacion() {
    showLoader('Cargando registros demostrativos de Demanda Inducida...');
    setTimeout(() => {
        const demoData = [
            { codigo: '890201', nombre: 'CONSULTA CONTROL SEGUIMIENTO EMSSANAR', finalidad: '10', cexterna: '38', centroprod: '1416', nombre_centroproduccion: 'PROGRAMA HIPERTENSION', nombre_prestador: 'C.S. PRIMERO DE MAYO', documento: '1143009988', factura: 'FEV-2001' },
            { codigo: '890202', nombre: 'CONSULTA PRIMERA VEZ CURSO VIDA ADULTO', finalidad: '11', cexterna: '38', centroprod: '1417', nombre_centroproduccion: 'CURSO DE VIDA', nombre_prestador: 'P.S. BRISAS DE MAYO', documento: '1144556677', factura: 'FEV-2002' }
        ];
        processData(demoData);
    }, 600);
}

// Export Excel with Formatting
async function exportExcelReport() {
    if (!state.inconsistencias.length) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Todas las inconsistencias');

    // Define Columns
    sheet.columns = [
        { header: 'CODIGO', key: 'codigo', width: 12 },
        { header: 'NOMBRE ACTIVIDAD', key: 'nombre', width: 35 },
        { header: 'FINALIDAD RIPS', key: 'finalidad', width: 15 },
        { header: 'CAUSA EXTERNA', key: 'cexterna', width: 15 },
        { header: 'IPS PRESTADOR', key: 'prestador', width: 28 },
        { header: 'DOCUMENTO', key: 'documento', width: 18 },
        { header: 'FACTURA', key: 'factura', width: 18 },
        { header: 'Inconsistencia a Corregir', key: 'inconsistencia', width: 55 }
    ];

    // Style Header Row (Yellow Background)
    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF00' }
        };
        cell.font = { bold: true, color: { argb: '000000' } };
    });

    // Add Data Rows (Red Font for Error)
    state.inconsistencias.forEach(item => {
        const row = sheet.addRow(item);
        const errorCell = row.getCell('inconsistencia');
        errorCell.font = { color: { argb: 'FF0000' }, bold: true };
    });

    // Generate File Download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Reporte_Inconsistencias_RFAST_${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
}

function showLoader(msg) {
    document.getElementById('loader-message').textContent = msg;
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}
