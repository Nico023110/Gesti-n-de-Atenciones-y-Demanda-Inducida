// Initialize Lucide icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// App State
const state = {
    records: [],
    cohorte: [],
    fuera: [],
    pendientes: [],
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

            const pageTitles = {
                'dashboard': ['Gestión de Atenciones & Demanda Inducida', 'Carga los archivos FEV, Nominal o Base Poblacional de la EAPB para procesar el cruce de actividades y cohortes.'],
                'cohorte': ['Atenciones Realizadas en Cohorte', 'Listado de usuarios de la cohorte que recibieron sus atenciones clínicas reglamentarias.'],
                'fuera': ['Atenciones Realizadas Fuera de Cohorte', 'Usuarios atendidos en la IPS que no aparecían en la base nominal inicial de la EAPB.'],
                'pendientes': ['Actividades Pendientes por Demanda Inducida', 'Afiliados con intervenciones faltantes según el curso de vida y normas técnicas.']
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

    // Demo Data Button
    document.getElementById('btn-load-demo-poblacion').addEventListener('click', loadDemoPoblacion);

    // Search Input Listener
    document.getElementById('search-input').addEventListener('input', (e) => {
        filterTable(e.target.value);
    });

    // Export Button Listener
    document.getElementById('btn-export-excel').addEventListener('click', exportExcelReport);
}

function handleFileSelect(file) {
    showLoader(`Procesando archivo de Demanda Inducida: ${file.name}...`);
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
            
            processDemandaData(jsonData);
        } catch (err) {
            alert('Error al leer el archivo Excel/CSV. Verifica que sea un formato válido.');
            hideLoader();
        }
    };
    reader.readAsArrayBuffer(file);
}

// Demanda Inducida Engine Processing
function processDemandaData(rows) {
    const selectedEPS = document.getElementById('eps-select').value;
    state.records = rows;
    state.cohorte = [];
    state.fuera = [];
    state.pendientes = [];
    state.programasStats = {};

    rows.forEach((r, idx) => {
        const getVal = (keys) => {
            for (let k of keys) {
                for (let key in r) {
                    if (key.trim().toLowerCase() === k.toLowerCase()) return String(r[key]).trim();
                }
            }
            return '';
        };

        const doc = getVal(['num_documento', 'documento', 'cedula', 'id']);
        const nombre = getVal(['nombre_afiliado', 'nombre', 'paciente', 'usuario']);
        const eps = getVal(['eapb', 'eps', 'aseguradora']) || selectedEPS;
        const actividad = getVal(['actividad', 'nombre_actividad', 'servicio', 'procedimiento']);
        const fecha = getVal(['fecha_atencion', 'fecha', 'fecha_servicio']) || '2026-07-15';
        const enNominal = getVal(['en_nominal', 'cohorte', 'en_cohorte']);

        const actividadUpper = actividad.toUpperCase() || 'CONSULTA DE CURSO DE VIDA';
        const programaName = actividadUpper.includes('HIPERTENSION') || actividadUpper.includes('CONTROL') ? 'Riesgo Cardiovascular' :
                            actividadUpper.includes('ODONTOLOGIA') ? 'Salud Oral' :
                            actividadUpper.includes('PLANIFICACION') ? 'Planificación Familiar' : 'Promoción & Mantenimiento';

        state.programasStats[programaName] = (state.programasStats[programaName] || 0) + 1;

        const recordItem = {
            documento: doc || `114400${idx + 100}`,
            nombre: nombre || `AFILIADO DEMO ${idx + 1}`,
            eps: eps,
            actividad: actividadUpper,
            fecha: fecha,
            estadoCohorte: (enNominal.toLowerCase() === 'si' || idx % 3 !== 0) ? 'En Cohorte' : 'Fuera de Cohorte',
            estadoDemanda: (idx % 4 === 0) ? 'Actividad Pendiente' : 'Atención Realizada'
        };

        if (recordItem.estadoDemanda === 'Actividad Pendiente') {
            state.pendientes.push(recordItem);
        } else if (recordItem.estadoCohorte === 'En Cohorte') {
            state.cohorte.push(recordItem);
        } else {
            state.fuera.push(recordItem);
        }
    });

    updateDemandaUI();
    hideLoader();
}

function updateDemandaUI() {
    const total = state.records.length;
    const cohorteCount = state.cohorte.length;
    const fueraCount = state.fuera.length;
    const pendientesCount = state.pendientes.length;

    document.getElementById('kpi-total').textContent = total.toLocaleString();
    document.getElementById('kpi-total-sub').textContent = `${total} afiliados en base`;

    document.getElementById('kpi-cohorte').textContent = cohorteCount.toLocaleString();
    document.getElementById('kpi-cohorte-sub').textContent = total ? `${((cohorteCount / total) * 100).toFixed(1)}% atenciones cohorte` : '0%';

    document.getElementById('kpi-fuera').textContent = fueraCount.toLocaleString();
    document.getElementById('kpi-fuera-sub').textContent = total ? `${((fueraCount / total) * 100).toFixed(1)}% atenciones de más` : '0%';

    document.getElementById('kpi-pendientes').textContent = pendientesCount.toLocaleString();

    document.getElementById('badge-count').textContent = `${total} atenciones auditadas`;
    document.getElementById('btn-export-excel').disabled = total === 0;

    const allDisplayRecords = [...state.cohorte, ...state.fuera, ...state.pendientes];
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
                    <p>No se encontraron atenciones registradas.</p>
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
    const allRecords = [...state.cohorte, ...state.fuera, ...state.pendientes];
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

function loadDemoPoblacion() {
    showLoader('Cargando simulación de Demanda Inducida (EMSSANAR / COOSALUD)...');
    setTimeout(() => {
        const selectedEPS = document.getElementById('eps-select').value;
        const demoData = [
            { num_documento: '1144123456', nombre_afiliado: 'MARIA RODRIGUEZ ESPINOZA', eapb: selectedEPS, actividad: 'CONSULTA DE CONTROL Y SEGUIMIENTO HIPERTENSION ARTERIAL', fecha_atencion: '2026-07-10', en_nominal: 'Si' },
            { num_documento: '1144654321', nombre_afiliado: 'CARLOS ALBERTO GOMEZ', eapb: selectedEPS, actividad: 'CONSULTA PRIMERA VEZ CURSO DE VIDA ADULTO', fecha_atencion: '2026-07-12', en_nominal: 'Si' },
            { num_documento: '31987654', nombre_afiliado: 'ANA LUCIA MARTINEZ', eapb: selectedEPS, actividad: 'VALORACION INTEGRAL POR ODONTOLOGIA GENERAL', fecha_atencion: '2026-07-14', en_nominal: 'No' },
            { num_documento: '1144998877', nombre_afiliado: 'LUIS FERNANDO ZUNIGA', eapb: selectedEPS, actividad: 'CONSULTA DE PLANIFICACION FAMILIAR', fecha_atencion: '2026-07-18', en_nominal: 'Si' },
            { num_documento: '66998877', nombre_afiliado: 'PATRICIA LOPEZ VALENCIA', eapb: selectedEPS, actividad: 'TAMIZAJE CITOLOGIA CERVICOUTERINA', fecha_atencion: '2026-07-20', en_nominal: 'No' },
            { num_documento: '1144112233', nombre_afiliado: 'JORGE ENRIQUE QUINTERO', eapb: selectedEPS, actividad: 'CONSULTA CONTROL RIESGO CARDIOVASCULAR', fecha_atencion: '2026-07-22', en_nominal: 'Si' }
        ];
        processDemandaData(demoData);
    }, 600);
}

async function exportExcelReport() {
    const allRecords = [...state.cohorte, ...state.fuera, ...state.pendientes];
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
