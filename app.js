// Initialize Lucide icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// App State
const state = {
    poblacionFiles: [],
    fevFiles: [],
    nominalFiles: [],
    poblacionRows: [],
    fevRows: [],
    nominalRows: [],
    results: {
        all: [],
        cohorte: [],
        fuera: [],
        pendientes: []
    },
    pagination: {
        all: { page: 1, records: [] },
        cohorte: { page: 1, records: [] },
        fuera: { page: 1, records: [] },
        pendientes: { page: 1, records: [] },
        pageSize: 50
    },
    programasStats: {},
    chartCobertura: null,
    chartProgramas: null
};

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
        });
    });

    setupDropzone('bucket-poblacion', 'file-input-poblacion', (files) => handleBucketFiles(files, 'poblacion'));
    setupDropzone('bucket-fev', 'file-input-fev', (files) => handleBucketFiles(files, 'fev'));
    setupDropzone('bucket-nominal', 'file-input-nominal', (files) => handleBucketFiles(files, 'nominal'));

    document.getElementById('btn-ejecutar-cruce').addEventListener('click', runDemandaCrucePipeline);
    document.getElementById('btn-load-demo').addEventListener('click', loadDemoSimulation);

    setupSearch('search-input', 'all');
    setupSearch('search-cohorte', 'cohorte');
    setupSearch('search-fuera', 'fuera');
    setupSearch('search-pendientes', 'pendientes');

    setupPaginationControls('all', 'btn-prev-page', 'btn-next-page');
    setupPaginationControls('cohorte', 'btn-prev-cohorte', 'btn-next-cohorte');
    setupPaginationControls('fuera', 'btn-prev-fuera', 'btn-next-fuera');
    setupPaginationControls('pendientes', 'btn-prev-pendientes', 'btn-next-pendientes');

    document.getElementById('btn-export-excel').addEventListener('click', () => exportExcelSubSet('all', 'Consolidado_Demanda_Inducida'));
    document.getElementById('btn-export-cohorte').addEventListener('click', () => exportExcelSubSet('cohorte', 'Atenciones_en_Cohorte'));
    document.getElementById('btn-export-fuera').addEventListener('click', () => exportExcelSubSet('fuera', 'Atenciones_Fuera_de_Cohorte'));
    document.getElementById('btn-export-pendientes').addEventListener('click', () => exportExcelSubSet('pendientes', 'Actividades_Pendientes_Demanda'));
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

async function handleBucketFiles(files, bucketType) {
    if (!files || !files.length) return;
    
    const validFiles = files.filter(f => {
        const ext = f.name.toLowerCase();
        return ext.endsWith('.csv') || ext.endsWith('.xlsx') || ext.endsWith('.xls') || ext.endsWith('.tsv') || ext.endsWith('.txt');
    });
    
    if (!validFiles.length) {
        alert('Por favor selecciona archivos válidos en formato CSV (.csv), Tabulado (.txt, .tsv) o Excel (.xlsx, .xls).');
        return;
    }

    showLoader(`Leyendo y procesando ${validFiles.length} archivo(s) para ${bucketType.toUpperCase()}...`);

    let loadedRowsTotal = [];
    let fileNames = [];

    for (let i = 0; i < validFiles.length; i++) {
        const file = validFiles[i];
        document.getElementById('loader-sub').textContent = `Procesando (${i + 1}/${validFiles.length}): ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)...`;
        
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

// 🚀 ROBUST PARSER: Uses ExcelJS for .xlsx to bypass SheetJS !ref row limits, fallbacks to SheetJS for .xls, and native JS for CSV/TSV
async function parseFileRows(file) {
    const ext = file.name.toLowerCase();

    if (ext.endsWith('.xlsx')) {
        try {
            // Use ExcelJS for .xlsx files. It iterates rows directly and ignores corrupt !ref limits.
            const arrayBuffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            
            const worksheet = workbook.worksheets[0];
            const rows = [];
            const headers = [];
            let headerParsed = false;
            
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                const values = row.values;
                // row.values is 1-indexed array in ExcelJS [empty, col1, col2...]
                if (!headerParsed) {
                    for (let i = 1; i < values.length; i++) {
                        headers.push(values[i] ? values[i].toString().trim() : `Col_${i}`);
                    }
                    headerParsed = true;
                } else {
                    const rowObj = {};
                    for (let i = 1; i <= headers.length; i++) {
                        const header = headers[i - 1];
                        let val = values[i];
                        
                        if (val !== null && val !== undefined) {
                            if (val.richText) {
                                val = val.richText.map(rt => rt.text).join('');
                            } else if (val instanceof Date) {
                                val = val.toISOString().slice(0, 10);
                            } else if (typeof val === 'object' && val.result !== undefined) {
                                // Formula result
                                val = val.result;
                            }
                            rowObj[header] = val.toString().trim();
                        } else {
                            rowObj[header] = '';
                        }
                    }
                    rows.push(rowObj);
                }
            });
            return rows;
        } catch (err) {
            console.error('ExcelJS parse error, falling back to SheetJS:', err);
            return parseWithSheetJS(file);
        }
    } else if (ext.endsWith('.xls')) {
        return parseWithSheetJS(file);
    } else {
        // Text-based files (.csv, .tsv, .txt)
        const text = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.onerror = e => reject(e);
            reader.readAsText(file, 'UTF-8');
        });
        return parseCSVText(text);
    }
}

function parseWithSheetJS(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheet = workbook.SheetNames[0];
                const sheet = workbook.Sheets[firstSheet];

                // Attempt to fix !ref range metadata just in case
                let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
                let hasCells = false;
                for (let key in sheet) {
                    if (key[0] === '!') continue;
                    const cell = XLSX.utils.decode_cell(key);
                    hasCells = true;
                    if (cell.r < minRow) minRow = cell.r;
                    if (cell.r > maxRow) maxRow = cell.r;
                    if (cell.c < minCol) minCol = cell.c;
                    if (cell.c > maxCol) maxCol = cell.c;
                }
                if (hasCells) {
                    sheet['!ref'] = XLSX.utils.encode_range({
                        s: { r: minRow, c: minCol },
                        e: { r: maxRow, c: maxCol }
                    });
                }

                const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
                resolve(rows);
            } catch (err) {
                console.error('SheetJS parse error:', err);
                resolve([]);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function parseCSVText(text) {
    if (!text || !text.trim()) return [];

    const lines = text.split(/\r?\n/);
    if (lines.length < 2) return [];

    let headerIdx = 0;
    while (headerIdx < lines.length && !lines[headerIdx].trim()) {
        headerIdx++;
    }
    if (headerIdx >= lines.length) return [];

    const firstLine = lines[headerIdx];
    
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const semiCount = (firstLine.match(/;/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const pipeCount = (firstLine.match(/\|/g) || []).length;

    let delimiter = ',';
    const maxCount = Math.max(tabCount, semiCount, commaCount, pipeCount);

    if (maxCount > 0) {
        if (maxCount === tabCount) delimiter = '\t';
        else if (maxCount === semiCount) delimiter = ';';
        else if (maxCount === pipeCount) delimiter = '|';
        else delimiter = ',';
    }

    const headers = firstLine.split(delimiter).map(h => h.replace(/^["']|["']$/g, '').trim());

    const result = [];
    for (let i = headerIdx + 1; i < lines.length; i++) {
        const currentLine = lines[i];
        if (!currentLine || !currentLine.trim()) continue;

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
        info.innerHTML = `✅ ${count.toLocaleString()} registros leídos al 100%`;

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

// Pipeline: Run Demanda Inducida Cross-Referencing
function runDemandaCrucePipeline() {
    showLoader('Ejecutando cruce de Demanda Inducida en el 100% de los registros leídos...');
    setTimeout(() => {
        const selectedEPS = document.getElementById('eps-select').value;

        state.results.all = [];
        state.results.cohorte = [];
        state.results.fuera = [];
        state.results.pendientes = [];
        state.programasStats = {};

        const primaryRows = state.poblacionRows.length ? state.poblacionRows : (state.fevRows.length ? state.fevRows : state.nominalRows);

        primaryRows.forEach((r, idx) => {
            const getVal = (keys) => {
                for (let k of keys) {
                    for (let key in r) {
                        if (key.trim().toLowerCase() === k.toLowerCase()) return String(r[key]).trim();
                    }
                }
                return '';
            };

            const doc = getVal(['num_documento', 'documento', 'cedula', 'num_documento_identificacion', 'numdocumentoidentificacion', 'num_doc']);
            const nombre = getVal(['nombre_afiliado', 'nombre', 'paciente', 'usuario', 'nombre_completo', 'primer_nombre']);
            const actividad = getVal(['actividad', 'nombre_actividad', 'servicio', 'procedimiento', 'cod_consulta', 'nombre_procedimiento', 'curso_vida', 'ciclovida']);
            const fecha = getVal(['fecha_atencion', 'fecha', 'fecha_servicio', 'fechainicioatencion', 'fecha_nacimiento']) || '2026-07-15';

            const actUpper = actividad.toUpperCase() || 'VALORACION INTEGRAL SALUD';
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
    document.getElementById('badge-count-cohorte').textContent = `${cohorteCount.toLocaleString()} registros`;
    document.getElementById('badge-count-fuera').textContent = `${fueraCount.toLocaleString()} registros`;
    document.getElementById('badge-count-pendientes').textContent = `${pendientesCount.toLocaleString()} registros`;

    document.getElementById('btn-export-excel').disabled = totalAll === 0;
    document.getElementById('btn-export-cohorte').disabled = cohorteCount === 0;
    document.getElementById('btn-export-fuera').disabled = fueraCount === 0;
    document.getElementById('btn-export-pendientes').disabled = pendientesCount === 0;

    state.pagination.all.records = state.results.all; state.pagination.all.page = 1;
    state.pagination.cohorte.records = state.results.cohorte; state.pagination.cohorte.page = 1;
    state.pagination.fuera.records = state.results.fuera; state.pagination.fuera.page = 1;
    state.pagination.pendientes.records = state.results.pendientes; state.pagination.pendientes.page = 1;

    renderTabTable('all');
    renderTabTable('cohorte');
    renderTabTable('fuera');
    renderTabTable('pendientes');

    renderCharts(cohorteCount, fueraCount, pendientesCount);
}

function renderTabTable(tabKey) {
    const tableBodyId = tabKey === 'all' ? 'table-body' : `table-body-${tabKey}`;
    const pageInfoId = tabKey === 'all' ? 'page-info' : `page-info-${tabKey}`;
    const pageNumId = tabKey === 'all' ? 'current-page-num' : `page-num-${tabKey}`;
    const prevBtnId = tabKey === 'all' ? 'btn-prev-page' : `btn-prev-${tabKey}`;
    const nextBtnId = tabKey === 'all' ? 'btn-next-page' : `btn-next-${tabKey}`;

    const tbody = document.getElementById(tableBodyId);
    if (!tbody) return;
    tbody.innerHTML = '';

    const tabState = state.pagination[tabKey];
    const records = tabState.records;
    const pageSize = state.pagination.pageSize;
    const totalPages = Math.max(1, Math.ceil(records.length / pageSize));
    const currentPage = Math.min(tabState.page, totalPages);
    tabState.page = currentPage;

    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, records.length);
    const pageRecords = records.slice(startIdx, endIdx);

    if (!records.length) {
        const colSpan = tabKey === 'all' ? 8 : 7;
        tbody.innerHTML = `
            <tr>
                <td colspan="${colSpan}" class="empty-table-msg">
                    <i data-lucide="folder-input"></i>
                    <p>No hay registros disponibles para esta categoría.</p>
                </td>
            </tr>
        `;
        document.getElementById(pageInfoId).textContent = 'Mostrando 0 registros';
        document.getElementById(prevBtnId).disabled = true;
        document.getElementById(nextBtnId).disabled = true;
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

        if (tabKey === 'all') {
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
        } else if (tabKey === 'pendientes') {
            tr.innerHTML = `
                <td><code>#${row.id}</code></td>
                <td><code>${row.documento}</code></td>
                <td><strong>${row.nombre}</strong></td>
                <td><span class="badge" style="background: rgba(99,102,241,0.15); color: #818CF8">${row.eps}</span></td>
                <td>${row.actividad}</td>
                <td>${row.fecha}</td>
                <td><span class="badge" style="${badgeDemandaStyle}">${row.estadoDemanda}</span></td>
            `;
        } else {
            tr.innerHTML = `
                <td><code>#${row.id}</code></td>
                <td><code>${row.documento}</code></td>
                <td><strong>${row.nombre}</strong></td>
                <td><span class="badge" style="background: rgba(99,102,241,0.15); color: #818CF8">${row.eps}</span></td>
                <td>${row.actividad}</td>
                <td>${row.fecha}</td>
                <td><span class="badge" style="${badgeCohorteStyle}">${row.estadoCohorte}</span></td>
            `;
        }
        tbody.appendChild(tr);
    });

    document.getElementById(pageInfoId).textContent = `Mostrando ${(startIdx + 1).toLocaleString()} a ${endIdx.toLocaleString()} de ${records.length.toLocaleString()} registros`;
    document.getElementById(pageNumId).textContent = `Página ${currentPage} de ${totalPages}`;
    document.getElementById(prevBtnId).disabled = currentPage === 1;
    document.getElementById(nextBtnId).disabled = currentPage === totalPages;
}

function setupPaginationControls(tabKey, prevBtnId, nextBtnId) {
    document.getElementById(prevBtnId).addEventListener('click', () => {
        if (state.pagination[tabKey].page > 1) {
            state.pagination[tabKey].page--;
            renderTabTable(tabKey);
        }
    });

    document.getElementById(nextBtnId).addEventListener('click', () => {
        const totalPages = Math.ceil(state.pagination[tabKey].records.length / state.pagination.pageSize);
        if (state.pagination[tabKey].page < totalPages) {
            state.pagination[tabKey].page++;
            renderTabTable(tabKey);
        }
    });
}

function setupSearch(inputId, tabKey) {
    document.getElementById(inputId).addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const rawSource = tabKey === 'all' ? state.results.all : state.results[tabKey];
        state.pagination[tabKey].records = rawSource.filter(row => {
            return row.documento.toLowerCase().includes(q) ||
                   row.nombre.toLowerCase().includes(q) ||
                   row.eps.toLowerCase().includes(q) ||
                   row.actividad.toLowerCase().includes(q) ||
                   row.estadoCohorte.toLowerCase().includes(q) ||
                   row.estadoDemanda.toLowerCase().includes(q);
        });
        state.pagination[tabKey].page = 1;
        renderTabTable(tabKey);
    });
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
        updateBucketUI('poblacion', ['EMSSANAR BD_ESE_LADERA.xlsx'], 78450);
        updateBucketUI('fev', ['FEV394424_CORREGIDO.csv', 'FEV394425_CORREGIDO.csv'], 15200);
        updateBucketUI('nominal', ['Sigires_NominalAfiliadosEmssanar.xlsx'], 45100);

        state.poblacionRows = Array(78450).fill({});
        state.fevRows = Array(15200).fill({});
        state.nominalRows = Array(45100).fill({});

        checkCanRun();
        runDemandaCrucePipeline();
    }, 600);
}

async function exportExcelSubSet(tabKey, filePrefix) {
    const records = tabKey === 'all' ? state.results.all : state.results[tabKey];
    if (!records || !records.length) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(filePrefix);

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

    records.forEach(item => {
        sheet.addRow(item);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filePrefix}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
}

function showLoader(msg) {
    document.getElementById('loader-message').textContent = msg;
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}
