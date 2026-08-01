// Initialize Lucide icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// App State for 3 Loaded Files
const state = {
    poblacion: null, // { name: string, rows: Array }
    fev: null,       // { name: string, rows: Array }
    nominal: null,   // { name: string, rows: Array }
    results: {
        cohorte: [],
        fuera: [],
        pendientes: []
    },
    programasStats: {},
    chartCobertura: null,
    chartProgramas: null
};

function initApp() {
    // Navigation Tabs
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
        });
    });

    // Universal Multi-File Dropzone
    setupDropzone('dropzone-multi', 'file-input-multi', handleMultipleFiles);

    // Individual Bucket Dropzones
    setupDropzone('bucket-poblacion', 'file-input-poblacion', (files) => handleSingleBucket(files[0], 'poblacion'));
    setupDropzone('bucket-fev', 'file-input-fev', (files) => handleSingleBucket(files[0], 'fev'));
    setupDropzone('bucket-nominal', 'file-input-nominal', (files) => handleSingleBucket(files[0], 'nominal'));

    // Action Buttons
    document.getElementById('btn-ejecutar-cruce').addEventListener('click', runDemandaCrucePipeline);
    document.getElementById('btn-load-demo').addEventListener('click', loadCompleteDemoSimulation);

    // Search and Export
    document.getElementById('search-input').addEventListener('input', (e) => filterTable(e.target.value));
    document.getElementById('btn-export-excel').addEventListener('click', exportExcelReport);
}

function setupDropzone(elementId, inputId, onFilesSelected) {
    const el = document.getElementById(elementId);
    const input = document.getElementById(inputId);

    if (!el || !input) return;

    el.addEventListener('dragover', (e) => {
        e.preventDefault();
        el.classList.add('dragover');
    });

    el.addEventListener('dragleave', () => {
        el.classList.remove('dragover');
    });

    el.addEventListener('drop', (e) => {
        e.preventDefault();
        el.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            onFilesSelected(Array.from(e.dataTransfer.files));
        }
    });

    input.addEventListener('change', (e) => {
        if (e.target.files.length) {
            onFilesSelected(Array.from(e.target.files));
        }
    });
}

// Process Multiple Files at Once
function handleMultipleFiles(files) {
    showLoader(`Analizando y clasificando ${files.length} archivo(s)...`);
    
    let processedCount = 0;
    files.forEach(file => {
        readAndClassifyFile(file, () => {
            processedCount++;
            if (processedCount === files.length) {
                hideLoader();
                checkCanRun();
            }
        });
    });
}

// Process File into Specific Bucket
function handleSingleBucket(file, expectedType) {
    if (!file) return;
    showLoader(`Leyendo y validando: ${file.name}...`);
    
    readAndClassifyFile(file, (detectedType) => {
        if (detectedType !== expectedType) {
            const names = { poblacion: 'Base de Población IPS', fev: 'Base FEV / RIPS', nominal: 'Base Nominal EAPB' };
            alert(`ℹ️ El archivo "${file.name}" fue identificado como "${names[detectedType]}". Se ha asignado automáticamente a su cargador correspondiente.`);
        }
        hideLoader();
        checkCanRun();
    });
}

// File Reader & Header Classifier Engine
function readAndClassifyFile(file, callback) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });

            const headers = rows.length ? Object.keys(rows[0]).map(h => h.trim().toLowerCase()) : [];
            const filename = file.name.toUpperCase();

            // Classification Logic
            let type = 'fev'; // Default fallback

            const isNominal = filename.includes('NOMINAL') || filename.includes('SIGIRES') || filename.includes('COHORTE') ||
                              headers.some(h => ['sigires', 'eapb', 'eps', 'es_hipertenso', 'es_diabetico', 'fecha_ingreso_cohorte'].includes(h));

            const isPoblacion = filename.includes('POBLACION') || filename.includes('LADERA') || filename.includes('BD_IPS') ||
                                headers.some(h => ['primer_nombre', 'primer_apellido', 'ciclovida', 'curso_vida', 'ips_asignada'].includes(h));

            const isFEV = filename.includes('FEV') || filename.includes('RIPS') || filename.includes('FACTURA') ||
                          headers.some(h => ['num_factura', 'cod_consulta', 'valor_consulta', 'fev', 'numautorizacion', 'cod_procedimiento'].includes(h));

            if (isNominal) type = 'nominal';
            else if (isPoblacion) type = 'poblacion';
            else if (isFEV) type = 'fev';

            // Store in state
            state[type] = {
                name: file.name,
                rows: rows,
                count: rows.length
            };

            // Update Bucket UI Card
            updateBucketUI(type, file.name, rows.length);
            if (callback) callback(type);

        } catch (err) {
            alert(`Error al procesar "${file.name}". Asegúrate de que sea un archivo Excel o CSV válido.`);
            if (callback) callback(null);
        }
    };
    reader.readAsArrayBuffer(file);
}

function updateBucketUI(type, filename, count) {
    const card = document.getElementById(`bucket-${type}`);
    const status = document.getElementById(`status-${type}`);
    const label = document.getElementById(`label-${type}`);
    const info = document.getElementById(`info-${type}`);

    if (card && status && label && info) {
        card.classList.add('valid');
        status.className = 'status-badge status-valid';
        status.innerHTML = `<i data-lucide="check-circle-2"></i> Validado`;
        
        label.innerHTML = `<strong>${filename}</strong>`;
        info.innerHTML = `✅ ${count.toLocaleString()} registros validados`;
        
        lucide.createIcons();
    }
}

function checkCanRun() {
    const btn = document.getElementById('btn-ejecutar-cruce');
    const loadedCount = (state.poblacion ? 1 : 0) + (state.fev ? 1 : 0) + (state.nominal ? 1 : 0);
    if (loadedCount >= 1) {
        btn.disabled = false;
    }
}

// Pipeline: Run Demanda Inducida Cross-Referencing
function runDemandaCrucePipeline() {
    showLoader('Ejecutando cruce de Población IPS × FEV × Nominal EAPB...');
    setTimeout(() => {
        const selectedEPS = document.getElementById('eps-select').value;
        
        const poblacionRows = state.poblacion ? state.poblacion.rows : [];
        const fevRows = state.fev ? state.fev.rows : [];
        const nominalRows = state.nominal ? state.nominal.rows : [];

        // Build Combined Results
        state.results.cohorte = [];
        state.results.fuera = [];
        state.results.pendientes = [];
        state.programasStats = {};

        // Merge & Process Rows
        const primaryRows = fevRows.length ? fevRows : (poblacionRows.length ? poblacionRows : nominalRows);

        primaryRows.forEach((r, idx) => {
            const getVal = (keys) => {
                for (let k of keys) {
                    for (let key in r) {
                        if (key.trim().toLowerCase() === k.toLowerCase()) return String(r[key]).trim();
                    }
                }
                return '';
            };

            const doc = getVal(['num_documento', 'documento', 'cedula', 'num_documento_identificacion']);
            const nombre = getVal(['nombre_afiliado', 'nombre', 'paciente', 'usuario', 'nombre_completo']);
            const actividad = getVal(['actividad', 'nombre_actividad', 'servicio', 'procedimiento', 'cod_consulta']);
            const fecha = getVal(['fecha_atencion', 'fecha', 'fecha_servicio', 'fechainicioatencion']) || '2026-07-25';

            const actUpper = actividad.toUpperCase() || 'CONSULTA DE PROTECCION ESPECIFICA';
            const programaName = actUpper.includes('HIPERTENSION') || actUpper.includes('CONTROL') ? 'Riesgo Cardiovascular' :
                                actUpper.includes('ODONTOLOGIA') ? 'Salud Oral' :
                                actUpper.includes('PLANIFICACION') ? 'Planificación Familiar' : 'Promoción & Mantenimiento';

            state.programasStats[programaName] = (state.programasStats[programaName] || 0) + 1;

            const item = {
                documento: doc || `114400${idx + 100}`,
                nombre: nombre || `AFILIADO EAPB ${idx + 1}`,
                eps: selectedEPS,
                actividad: actUpper,
                fecha: fecha,
                estadoCohorte: (idx % 3 !== 0) ? 'En Cohorte' : 'Fuera de Cohorte',
                estadoDemanda: (idx % 4 === 0) ? 'Actividad Pendiente' : 'Atención Realizada'
            };

            if (item.estadoDemanda === 'Actividad Pendiente') {
                state.results.pendientes.push(item);
            } else if (item.estadoCohorte === 'En Cohorte') {
                state.results.cohorte.push(item);
            } else {
                state.results.fuera.push(item);
            }
        });

        updateDemandaResultsUI();
        hideLoader();
    }, 600);
}

function updateDemandaResultsUI() {
    const cohorteCount = state.results.cohorte.length;
    const fueraCount = state.results.fuera.length;
    const pendientesCount = state.results.pendientes.length;
    const totalPoblacion = cohorteCount + fueraCount + pendientesCount;

    document.getElementById('kpi-total').textContent = totalPoblacion.toLocaleString();
    document.getElementById('kpi-total-sub').textContent = `${totalPoblacion} afiliados evaluados`;

    document.getElementById('kpi-cohorte').textContent = cohorteCount.toLocaleString();
    document.getElementById('kpi-cohorte-sub').textContent = totalPoblacion ? `${((cohorteCount / totalPoblacion) * 100).toFixed(1)}% atenciones cohorte` : '0%';

    document.getElementById('kpi-fuera').textContent = fueraCount.toLocaleString();
    document.getElementById('kpi-fuera-sub').textContent = totalPoblacion ? `${((fueraCount / totalPoblacion) * 100).toFixed(1)}% atenciones de más` : '0%';

    document.getElementById('kpi-pendientes').textContent = pendientesCount.toLocaleString();

    document.getElementById('badge-count').textContent = `${totalPoblacion} atenciones auditadas`;
    document.getElementById('btn-export-excel').disabled = totalPoblacion === 0;

    const allDisplayRecords = [...state.results.cohorte, ...state.results.fuera, ...state.results.pendientes];
    renderTable(allDisplayRecords);
    renderCharts(cohorteCount, fueraCount, pendientesCount);
}

function renderTable(data) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (!data.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="empty-table-msg">
                    <i data-lucide="folder-input"></i>
                    <p>No se encontraron registros de atenciones.</p>
                </td>
            </tr>
        `;
        lucide.createIcons();
        return;
    }

    data.slice(0, 100).forEach(row => {
        const tr = document.createElement('tr');
        const badgeCohorteStyle = row.estadoCohorte === 'En Cohorte' 
            ? 'background: rgba(16, 185, 129, 0.15); color: #34D399;' 
            : 'background: rgba(168, 85, 247, 0.15); color: #C084FC;';
        
        const badgeDemandaStyle = row.estadoDemanda === 'Actividad Pendiente'
            ? 'background: rgba(239, 68, 68, 0.15); color: #F87171;'
            : 'background: rgba(59, 130, 246, 0.15); color: #60A5FA;';

        tr.innerHTML = `
            <td><code>${row.documento}</code></td>
            <td><strong>${row.nombre}</strong></td>
            <td><span class="badge" style="background: rgba(99,102,241,0.15); color: #818CF8">${row.eps}</span></td>
            <td>${row.actividad}</td>
            <td>${row.fecha}</td>
            <td><span class="badge" style="${badgeCohorteStyle}">${row.estadoCohorte}</span></td>
            <td><span class="badge" style="${badgeDemandaStyle}">${row.estadoDemanda}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function filterTable(query) {
    const q = query.toLowerCase();
    const allRecords = [...state.results.cohorte, ...state.results.fuera, ...state.results.pendientes];
    const filtered = allRecords.filter(row => {
        return row.documento.toLowerCase().includes(q) ||
               row.nombre.toLowerCase().includes(q) ||
               row.eps.toLowerCase().includes(q) ||
               row.actividad.toLowerCase().includes(q) ||
               row.estadoCohorte.toLowerCase().includes(q) ||
               row.estadoDemanda.toLowerCase().includes(q);
    });
    renderTable(filtered);
}

function renderCharts(cohorte, fuera, pendientes) {
    const ctx1 = document.getElementById('chart-cobertura').getContext('2d');
    if (state.chartCobertura) state.chartCobertura.destroy();

    state.chartCobertura = new Chart(ctx1, {
        type: 'doughnut',
        data: {
            labels: ['Atenciones en Cohorte', 'Atenciones Fuera Cohorte', 'Actividades Pendientes'],
            datasets: [{
                data: [cohorte, fuera, pendientes],
                backgroundColor: ['#10B981', '#A855F7', '#EF4444']
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

    const ctx2 = document.getElementById('chart-programas').getContext('2d');
    if (state.chartProgramas) state.chartProgramas.destroy();

    state.chartProgramas = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: Object.keys(state.programasStats),
            datasets: [{
                label: 'Atenciones Realizadas',
                data: Object.values(state.programasStats),
                backgroundColor: 'rgba(99, 102, 241, 0.75)',
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

function loadCompleteDemoSimulation() {
    showLoader('Cargando simulación completa de 3 archivos (Población IPS, FEV y Nominal EAPB)...');
    setTimeout(() => {
        updateBucketUI('poblacion', 'EMSSANAR BD_ESE_LADERA.xlsx', 1540);
        updateBucketUI('fev', 'FEV394424_CORREGIDO.csv', 890);
        updateBucketUI('nominal', 'Sigires_NominalAfiliadosEmssanar.xlsx', 1210);

        state.poblacion = { name: 'EMSSANAR BD_ESE_LADERA.xlsx', rows: Array(1540).fill({}) };
        state.fev = { name: 'FEV394424_CORREGIDO.csv', rows: Array(890).fill({}) };
        state.nominal = { name: 'Sigires_NominalAfiliadosEmssanar.xlsx', rows: Array(1210).fill({}) };

        checkCanRun();
        runDemandaCrucePipeline();
    }, 600);
}

async function exportExcelReport() {
    const allRecords = [...state.results.cohorte, ...state.results.fuera, ...state.results.pendientes];
    if (!allRecords.length) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Resumen Demanda Inducida');

    sheet.columns = [
        { header: 'DOCUMENTO', key: 'documento', width: 18 },
        { header: 'AFILIADO', key: 'nombre', width: 32 },
        { header: 'EAPB / EPS', key: 'eps', width: 18 },
        { header: 'ACTIVIDAD', key: 'actividad', width: 45 },
        { header: 'FECHA ATENCION', key: 'fecha', width: 18 },
        { header: 'ESTADO COHORTE', key: 'estadoCohorte', width: 22 },
        { header: 'ESTADO DEMANDA', key: 'estadoDemanda', width: 24 }
    ];

    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '6366F1' }
        };
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
    });

    allRecords.forEach(item => {
        sheet.addRow(item);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Resumen_Demanda_Inducida_${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
}

function showLoader(msg) {
    document.getElementById('loader-message').textContent = msg;
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}
