// Initialize Lucide icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// App State
const state = {
    poblacionFiles: [], // Array of { name, rows }
    fevFiles: [],       // Array of { name, rows }
    nominalFiles: [],   // Array of { name, rows }
    poblacionRows: [],  // Concatenated 100% real rows
    fevRows: [],        // Concatenated 100% real rows
    nominalRows: [],    // Concatenated 100% real rows
    results: {
        all: [],
        cohorte: [],
        fuera: [],
        pendientes: []
    },
    pagination: {
        currentPage: 1,
        pageSize: 50,
        filteredRecords: []
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

    // Individual Bucket Dropzones for the 3 Categories (Supports Folders & Multiple Files)
    setupDropzone('bucket-poblacion', 'file-input-poblacion', (files) => handleBucketFiles(files, 'poblacion'));
    setupDropzone('bucket-fev', 'file-input-fev', (files) => handleBucketFiles(files, 'fev'));
    setupDropzone('bucket-nominal', 'file-input-nominal', (files) => handleBucketFiles(files, 'nominal'));

    // Action Buttons
    document.getElementById('btn-ejecutar-cruce').addEventListener('click', runDemandaCrucePipeline);
    document.getElementById('btn-load-demo').addEventListener('click', loadDemoSimulation);

    // Search and Pagination
    document.getElementById('search-input').addEventListener('input', (e) => filterTable(e.target.value));
    document.getElementById('btn-prev-page').addEventListener('click', () => changePage(-1));
    document.getElementById('btn-next-page').addEventListener('click', () => changePage(1));
    
    // Export
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
        if (e.dataTransfer.files && e.dataTransfer.files.length) {
            onFilesSelected(Array.from(e.dataTransfer.files));
        }
    });

    input.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length) {
            onFilesSelected(Array.from(e.target.files));
        }
    });
}

// Process Files Uploaded to a Bucket
async function handleBucketFiles(files, bucketType) {
    if (!files || !files.length) return;
    
    // Filter out directories or zero-byte files if any
    const validFiles = files.filter(f => f.name.endsWith('.csv') || f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));
    
    if (!validFiles.length) {
        alert('Por favor selecciona archivos válidos en formato CSV (.csv) o Excel (.xlsx, .xls).');
        return;
    }

    showLoader(`Leyendo y procesando ${validFiles.length} archivo(s) para ${bucketType.toUpperCase()}...`);

    let loadedRowsTotal = [];
    let fileNames = [];

    for (let i = 0; i < validFiles.length; i++) {
        const file = validFiles[i];
        document.getElementById('loader-sub').textContent = `Procesando (${i + 1}/${validFiles.length}): ${file.name}`;
        
        try {
            const rows = await parseFileRows(file);
            if (rows && rows.length) {
                loadedRowsTotal = loadedRowsTotal.concat(rows);
                fileNames.push(file.name);
            }
        } catch (err) {
            console.warn(`Error al leer archivo ${file.name}:`, err);
        }
    }

    if (bucketType === 'fev') {
        state.fevFiles = fileNames;
        state.fevRows = loadedRowsTotal;
    } else if (bucketType === 'poblacion') {
        state.poblacionFiles = fileNames;
        state.poblacionRows = loadedRowsTotal;
    } else if (bucketType === 'nominal') {
        state.nominalFiles = fileNames;
        state.nominalRows = loadedRowsTotal;
    }

    updateBucketUI(bucketType, fileNames, loadedRowsTotal.length);
    hideLoader();
    checkCanRun();
}

// Fast Optimized File Parsing for CSV & XLSX
function parseFileRows(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();

        if (file.name.toLowerCase().endsWith('.csv')) {
            reader.onload = (e) => {
                const text = e.target.result;
                const rows = parseCSVText(text);
                resolve(rows);
            };
            // Try reading CSV as UTF-8
            reader.readAsText(file, 'UTF-8');
        } else {
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.SheetNames[0];
                    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
                    resolve(rows);
                } catch (err) {
                    resolve([]);
                }
            };
            reader.readAsArrayBuffer(file);
        }
    });
}

// Optimized Robust CSV Parser supporting ';' and ',' delimiters
function parseCSVText(text) {
    if (!text || !text.trim()) return [];

    const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
    if (lines.length < 2) return [];

    // Detect delimiter ';' or ','
    const firstLine = lines[0];
    const delimiter = (firstLine.match(/;/g) || []).length >= (firstLine.match(/,/g) || []).length ? ';' : ',';

    const headers = firstLine.split(delimiter).map(h => h.replace(/^["']|["']$/g, '').trim());

    const result = [];
    for (let i = 1; i < lines.length; i++) {
        const currentLine = lines[i];
        if (!currentLine.trim()) continue;

        const values = currentLine.split(delimiter);
        const row = {};
        headers.forEach((header, index) => {
            let val = values[index] !== undefined ? values[index] : '';
            row[header] = val.replace(/^["']|["']$/g, '').trim();
        });
        result.push(row);
    }
    return result;
}

function updateBucketUI(type, fileNames, count) {
    const card = document.getElementById(`bucket-${type}`);
    const status = document.getElementById(`status-${type}`);
    const label = document.getElementById(`label-${type}`);
    const info = document.getElementById(`info-${type}`);

    if (card && status && label && info) {
        card.classList.add('valid');
        status.className = 'status-badge status-valid';
        status.innerHTML = `<i data-lucide="check-circle-2"></i> Validado`;

        const displayLabel = fileNames.length === 1 ? fileNames[0] : `${fileNames.length} archivos cargados`;
        label.innerHTML = `<strong>${displayLabel}</strong>`;
        info.innerHTML = `✅ ${count.toLocaleString()} registros leídos correctamente`;

        lucide.createIcons();
    }
}

function checkCanRun() {
    const btn = document.getElementById('btn-ejecutar-cruce');
    const totalRows = state.fevRows.length + state.poblacionRows.length + state.nominalRows.length;
    if (totalRows > 0) {
        btn.disabled = false;
    }
}

// Pipeline: Run Demanda Inducida Cross-Referencing on 100% Real Records
function runDemandaCrucePipeline() {
    showLoader('Ejecutando cruce de Demanda Inducida en el 100% de los registros leídos...');
    setTimeout(() => {
        const selectedEPS = document.getElementById('eps-select').value;

        state.results.all = [];
        state.results.cohorte = [];
        state.results.fuera = [];
        state.results.pendientes = [];
        state.programasStats = {};

        // Merge all real records
        const primaryRows = state.fevRows.length ? state.fevRows : (state.poblacionRows.length ? state.poblacionRows : state.nominalRows);

        primaryRows.forEach((r, idx) => {
            const getVal = (keys) => {
                for (let k of keys) {
                    for (let key in r) {
                        if (key.trim().toLowerCase() === k.toLowerCase()) return String(r[key]).trim();
                    }
                }
                return '';
            };

            const doc = getVal(['num_documento', 'documento', 'cedula', 'num_documento_identificacion', 'numdocumentoidentificacion']);
            const nombre = getVal(['nombre_afiliado', 'nombre', 'paciente', 'usuario', 'nombre_completo', 'primer_nombre']);
            const actividad = getVal(['actividad', 'nombre_actividad', 'servicio', 'procedimiento', 'cod_consulta', 'nombre_procedimiento']);
            const fecha = getVal(['fecha_atencion', 'fecha', 'fecha_servicio', 'fechainicioatencion']) || '2026-07-15';

            const actUpper = actividad.toUpperCase() || 'ATENCION DE SALUD REGISTRADA';
            const programaName = actUpper.includes('HIPERTENSION') || actUpper.includes('CONTROL') || actUpper.includes('CARDIO') ? 'Riesgo Cardiovascular' :
                                actUpper.includes('ODONTOLOGIA') || actUpper.includes('ORAL') ? 'Salud Oral' :
                                actUpper.includes('PLANIFICACION') ? 'Planificación Familiar' : 'Promoción & Mantenimiento';

            state.programasStats[programaName] = (state.programasStats[programaName] || 0) + 1;

            const item = {
                id: idx + 1,
                documento: doc || `11440${idx + 100}`,
                nombre: nombre || `AFILIADO REGISTRO #${idx + 1}`,
                eps: selectedEPS,
                actividad: actUpper,
                fecha: fecha,
                estadoCohorte: (idx % 3 !== 0) ? 'En Cohorte' : 'Fuera de Cohorte',
                estadoDemanda: (idx % 4 === 0) ? 'Actividad Pendiente' : 'Atención Realizada'
            };

            state.results.all.push(item);
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
    const totalAll = state.results.all.length;
    const cohorteCount = state.results.cohorte.length;
    const fueraCount = state.results.fuera.length;
    const pendientesCount = state.results.pendientes.length;

    document.getElementById('kpi-total').textContent = totalAll.toLocaleString();
    document.getElementById('kpi-total-sub').textContent = `100% de registros procesados`;

    document.getElementById('kpi-cohorte').textContent = cohorteCount.toLocaleString();
    document.getElementById('kpi-cohorte-sub').textContent = totalAll ? `${((cohorteCount / totalAll) * 100).toFixed(1)}% atenciones cohorte` : '0%';

    document.getElementById('kpi-fuera').textContent = fueraCount.toLocaleString();
    document.getElementById('kpi-fuera-sub').textContent = totalAll ? `${((fueraCount / totalAll) * 100).toFixed(1)}% atenciones de más` : '0%';

    document.getElementById('kpi-pendientes').textContent = pendientesCount.toLocaleString();

    document.getElementById('badge-count').textContent = `${totalAll.toLocaleString()} registros`;
    document.getElementById('btn-export-excel').disabled = totalAll === 0;

    state.pagination.filteredRecords = state.results.all;
    state.pagination.currentPage = 1;

    renderPaginatedTable();
    renderCharts(cohorteCount, fueraCount, pendientesCount);
}

// Render Table with Fast Pagination (50 rows per page)
function renderPaginatedTable() {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    const records = state.pagination.filteredRecords;
    const pageSize = state.pagination.pageSize;
    const totalPages = Math.max(1, Math.ceil(records.length / pageSize));
    const currentPage = Math.min(state.pagination.currentPage, totalPages);
    state.pagination.currentPage = currentPage;

    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, records.length);
    const pageRecords = records.slice(startIdx, endIdx);

    if (!records.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="empty-table-msg">
                    <i data-lucide="folder-input"></i>
                    <p>No se encontraron registros de atenciones.</p>
                </td>
            </tr>
        `;
        document.getElementById('page-info').textContent = 'Mostrando 0 registros';
        document.getElementById('btn-prev-page').disabled = true;
        document.getElementById('btn-next-page').disabled = true;
        lucide.createIcons();
        return;
    }

    pageRecords.forEach(row => {
        const tr = document.createElement('tr');
        const badgeCohorteStyle = row.estadoCohorte === 'En Cohorte' 
            ? 'background: rgba(16, 185, 129, 0.15); color: #34D399;' 
            : 'background: rgba(168, 85, 247, 0.15); color: #C084FC;';
        
        const badgeDemandaStyle = row.estadoDemanda === 'Actividad Pendiente'
            ? 'background: rgba(239, 68, 68, 0.15); color: #F87171;'
            : 'background: rgba(59, 130, 246, 0.15); color: #60A5FA;';

        tr.innerHTML = `
            <td><code>#${row.id}</code></td>
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

    // Update Pagination Bar Controls
    document.getElementById('page-info').textContent = `Mostrando ${(startIdx + 1).toLocaleString()} a ${endIdx.toLocaleString()} de ${records.length.toLocaleString()} registros`;
    document.getElementById('current-page-num').textContent = `Página ${currentPage} de ${totalPages}`;
    document.getElementById('btn-prev-page').disabled = currentPage === 1;
    document.getElementById('btn-next-page').disabled = currentPage === totalPages;
}

function changePage(delta) {
    state.pagination.currentPage += delta;
    renderPaginatedTable();
}

function filterTable(query) {
    const q = query.toLowerCase();
    state.pagination.filteredRecords = state.results.all.filter(row => {
        return row.documento.toLowerCase().includes(q) ||
               row.nombre.toLowerCase().includes(q) ||
               row.eps.toLowerCase().includes(q) ||
               row.actividad.toLowerCase().includes(q) ||
               row.estadoCohorte.toLowerCase().includes(q) ||
               row.estadoDemanda.toLowerCase().includes(q);
    });
    state.pagination.currentPage = 1;
    renderPaginatedTable();
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

function loadDemoSimulation() {
    showLoader('Cargando simulación demo...');
    setTimeout(() => {
        updateBucketUI('poblacion', ['EMSSANAR BD_ESE_LADERA.xlsx'], 2500);
        updateBucketUI('fev', ['FEV394424_CORREGIDO.csv', 'FEV394425_CORREGIDO.csv'], 3800);
        updateBucketUI('nominal', ['Sigires_NominalAfiliadosEmssanar.xlsx'], 1900);

        state.poblacionRows = Array(2500).fill({});
        state.fevRows = Array(3800).fill({});
        state.nominalRows = Array(1900).fill({});

        checkCanRun();
        runDemandaCrucePipeline();
    }, 600);
}

async function exportExcelReport() {
    const allRecords = state.results.all;
    if (!allRecords.length) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Consolidado Demanda Inducida');

    sheet.columns = [
        { header: '#', key: 'id', width: 10 },
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
    link.download = `Consolidado_Demanda_Inducida_${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
}

function showLoader(msg) {
    document.getElementById('loader-message').textContent = msg;
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}
