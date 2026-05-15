// Elementos del DOM
const dropzoneVentas = document.getElementById('dropzone-ventas');
const fileVentas = document.getElementById('file-ventas');
const statusVentas = document.getElementById('status-ventas');

const dropzonePromos = document.getElementById('dropzone-promos');
const filePromos = document.getElementById('file-promos');
const statusPromos = document.getElementById('status-promos');

const btnAnalyze = document.getElementById('btn-analyze');
const btnReset = document.getElementById('btn-reset');
const uploadSection = document.getElementById('upload-section');
const loader = document.getElementById('loader');
const dashboardSection = document.getElementById('dashboard-section');

// LLM Elements
const llmBadge = document.getElementById('llm-status-badge');
const aiContainer = document.getElementById('ai-typing-container');

// Settings Elements
const groupBtns = document.querySelectorAll('.control-btn[data-group]');
const chartBtns = document.querySelectorAll('.control-btn[data-chart]');
const timelineTitle = document.getElementById('timeline-chart-title');

// Datos Globales y Estado
let dataVentas = null;
let dataPromos = null;
let charts = [];
let state = {
    groupBy: 'day', // 'day' | 'month'
    chartType: 'line' // 'line' | 'bar'
};

// --- LLM WORKER SETUP ---
const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });

worker.addEventListener('message', (event) => {
    const { status, message, output } = event.data;

    switch (status) {
        case 'loading':
            llmBadge.classList.remove('hidden');
            llmBadge.innerHTML = `<span class="pulse-dot"></span> Cargando cerebro de IA (${message})`;
            break;
        case 'ready':
            llmBadge.classList.remove('hidden');
            llmBadge.classList.add('ready');
            llmBadge.innerHTML = `✅ IA Lista (100% Local)`;
            break;
        case 'generating':
            aiContainer.innerHTML = '<p class="ai-placeholder">La Inteligencia Artificial está analizando los datos...</p>';
            break;
        case 'complete':
            typeText(aiContainer, output);
            break;
        case 'error':
            aiContainer.innerHTML = `<p style="color: #ef4444">Error en la IA: ${message}</p>`;
            break;
    }
});

function typeText(element, text) {
    element.innerHTML = '';
    let i = 0;
    const interval = setInterval(() => {
        if (i < text.length) {
            element.innerHTML += text.charAt(i);
            i++;
        } else {
            clearInterval(interval);
        }
    }, 15);
}

// --- DRAG & DROP SETUP ---
function setupDropzone(dropzone, input, statusElem, type) {
    dropzone.addEventListener('click', () => input.click());
    dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0], statusElem, type);
    });
    input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0], statusElem, type);
    });
}

setupDropzone(dropzoneVentas, fileVentas, statusVentas, 'ventas');
setupDropzone(dropzonePromos, filePromos, statusPromos, 'promos');

function handleFile(file, statusElem, type) {
    statusElem.textContent = "Leyendo archivo...";
    statusElem.className = "file-status";
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(firstSheet);
            
            if (type === 'ventas') dataVentas = json;
            else dataPromos = json;
            
            statusElem.textContent = `✅ ${file.name} cargado`;
            statusElem.className = "file-status loaded";
            
            if (dataVentas && dataPromos) btnAnalyze.disabled = false;
        } catch (error) {
            statusElem.textContent = "❌ Error al leer el archivo";
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function findColumn(row, keywords) {
    const keys = Object.keys(row);
    for (let key of keys) {
        const lowerKey = key.toLowerCase();
        if (keywords.some(kw => lowerKey.includes(kw))) return key;
    }
    return null;
}

// --- CONTROLES INTERACTIVOS ---
groupBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        groupBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.groupBy = btn.getAttribute('data-group');
        timelineTitle.textContent = state.groupBy === 'month' ? 'Ventas Promedio Mensuales' : 'Comportamiento de Ventas Diarias';
        processData(false); // No regenerar IA si solo cambiamos la vista, o sí? Dejémoslo solo visual para no tardar
    });
});

chartBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        chartBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.chartType = btn.getAttribute('data-chart');
        processData(false);
    });
});

btnAnalyze.addEventListener('click', () => {
    uploadSection.classList.add('hidden');
    loader.classList.remove('hidden');
    setTimeout(() => { processData(true); }, 500);
});

btnReset.addEventListener('click', () => {
    dashboardSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    dataVentas = null; dataPromos = null;
    statusVentas.textContent = "Esperando archivo..."; statusVentas.className = "file-status";
    statusPromos.textContent = "Esperando archivo..."; statusPromos.className = "file-status";
    btnAnalyze.disabled = true;
    charts.forEach(c => c.destroy());
    charts = [];
    aiContainer.innerHTML = '<p class="ai-placeholder">Esperando a la IA...</p>';
});

// --- LÓGICA PRINCIPAL ---
function processData(generateAI = false) {
    try {
        if (!dataVentas || dataVentas.length === 0) throw new Error("Base vacía");
        
        const vRow = dataVentas[0];
        const colFecha = findColumn(vRow, ['fecha', 'date', 'dia']);
        const colProducto = findColumn(vRow, ['producto', 'product', 'item', 'sku']);
        const colVenta = findColumn(vRow, ['venta', 'sales', 'monto', 'ingreso', 'revenue', 'cantidad']);
        
        const pRow = dataPromos[0] || {};
        const colPromoName = findColumn(pRow, ['promo', 'campaña', 'campaign', 'nombre']);
        const colPromoProd = findColumn(pRow, ['producto', 'product', 'item', 'sku']);
        
        const promoMap = {};
        dataPromos.forEach(p => {
            const prod = String(p[colPromoProd || colProducto] || '').trim().toLowerCase();
            const promo = p[colPromoName] || 'Sí';
            if (prod) promoMap[prod] = promo;
        });

        let totalSales = 0;
        let promoSales = 0;
        const timelineData = {}; // Para agrupar
        const productData = {};
        const promoPerformance = {};

        dataVentas.forEach(v => {
            const val = parseFloat(v[colVenta]) || 0;
            const prod = String(v[colProducto] || 'Desconocido').trim();
            const rawDate = String(v[colFecha] || 'Sin Fecha').trim();
            
            // Agrupación de fechas
            let dateKey = rawDate;
            if (state.groupBy === 'month' && rawDate.length >= 7) {
                // Si viene como Excel serial, habría que convertirlo, pero asumiremos string YYYY-MM-DD
                // Intentar extraer YYYY-MM
                const match = rawDate.match(/^(\d{4}[-/]\d{2})/);
                if (match) dateKey = match[1].replace('/', '-');
                else dateKey = rawDate; // Fallback
            }

            const prodLower = prod.toLowerCase();
            const hasPromo = promoMap[prodLower];
            const promoName = hasPromo ? promoMap[prodLower] : 'Sin Promoción';

            totalSales += val;
            if (hasPromo) promoSales += val;

            // Timeline Data
            if (!timelineData[dateKey]) timelineData[dateKey] = { normal: 0, promo: 0, count: 0 };
            if (hasPromo) timelineData[dateKey].promo += val;
            else timelineData[dateKey].normal += val;
            timelineData[dateKey].count += 1;

            // Product Data
            if (!productData[prod]) productData[prod] = { normal: 0, promo: 0 };
            if (hasPromo) productData[prod].promo += val;
            else productData[prod].normal += val;

            // Promo Data
            if (hasPromo) {
                if (!promoPerformance[promoName]) promoPerformance[promoName] = 0;
                promoPerformance[promoName] += val;
            }
        });

        // Si agrupa por mes, calcular el promedio en lugar de la suma total del mes?
        // El usuario pidió "comportamientos de las ventas en promedio por mes".
        // Transformar la data del timeline a promedio si state.groupBy === 'month'
        const chartTimelineData = {};
        Object.keys(timelineData).forEach(key => {
            if (state.groupBy === 'month') {
                // Promedio de venta por registro en ese mes (o promedio diario si tuviéramos días únicos)
                // Para simplificar, dividiremos por la cantidad de registros (días).
                const count = timelineData[key].count || 1;
                chartTimelineData[key] = {
                    promo: timelineData[key].promo / count,
                    normal: timelineData[key].normal / count
                };
            } else {
                chartTimelineData[key] = timelineData[key];
            }
        });

        // KPIs
        document.getElementById('kpi-total-sales').textContent = `$${totalSales.toLocaleString('es-MX', {maximumFractionDigits:0})}`;
        document.getElementById('kpi-promo-sales').textContent = `$${promoSales.toLocaleString('es-MX', {maximumFractionDigits:0})}`;
        const promoPct = totalSales > 0 ? ((promoSales / totalSales) * 100).toFixed(1) : 0;
        document.getElementById('kpi-promo-percentage').textContent = `${promoPct}% del total`;

        const normalSales = totalSales - promoSales;
        const lift = normalSales > 0 ? ((promoSales / normalSales) * 100).toFixed(1) : 0;
        document.getElementById('kpi-lift').textContent = `+${lift}%`;

        // Destruir gráficas viejas
        charts.forEach(c => c.destroy());
        charts = [];

        // Render
        renderTimelineChart(chartTimelineData);
        renderProductChart(productData);
        renderPromoChart(promoPerformance);

        // Llamar a la IA
        if (generateAI) {
            let bestPromo = "N/A";
            if (Object.keys(promoPerformance).length > 0) {
                bestPromo = Object.keys(promoPerformance).reduce((a, b) => promoPerformance[a] > promoPerformance[b] ? a : b);
            }
            
            worker.postMessage({
                action: 'analyze',
                stats: { totalSales, promoSales, promoPct, lift, bestPromo }
            });
        }

        loader.classList.add('hidden');
        dashboardSection.classList.remove('hidden');

    } catch (err) {
        alert("Error procesando datos: " + err.message);
        loader.classList.add('hidden');
        uploadSection.classList.remove('hidden');
    }
}

// --- GRÁFICAS ---
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = 'Outfit';

function renderTimelineChart(data) {
    const ctx = document.getElementById('timelineChart').getContext('2d');
    const labels = Object.keys(data).sort();
    const promoData = labels.map(l => data[l].promo);
    const normalData = labels.map(l => data[l].normal);

    const isLine = state.chartType === 'line';

    const chart = new Chart(ctx, {
        type: isLine ? 'line' : 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: state.groupBy === 'month' ? 'Promedio Promo/Registro' : 'Ventas en Promoción',
                    data: promoData,
                    borderColor: '#8b5cf6',
                    backgroundColor: isLine ? 'rgba(139, 92, 246, 0.2)' : '#8b5cf6',
                    fill: isLine,
                    tension: 0.4
                },
                {
                    label: state.groupBy === 'month' ? 'Promedio Normal/Registro' : 'Ventas Normales',
                    data: normalData,
                    borderColor: '#3b82f6',
                    backgroundColor: isLine ? 'transparent' : '#3b82f6',
                    tension: 0.4
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
                x: { stacked: !isLine, grid: { display: false } }
            }
        }
    });
    charts.push(chart);
}

function renderProductChart(data) {
    const ctx = document.getElementById('productChart').getContext('2d');
    const sortedProds = Object.keys(data).sort((a,b) => (data[b].promo + data[b].normal) - (data[a].promo + data[a].normal)).slice(0, 7);
    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sortedProds,
            datasets: [
                { label: 'Con Promo', data: sortedProds.map(p => data[p].promo), backgroundColor: '#8b5cf6', borderRadius: 4 },
                { label: 'Sin Promo', data: sortedProds.map(p => data[p].normal), backgroundColor: '#3b82f6', borderRadius: 4 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, grid: { color: 'rgba(255,255,255,0.05)' } } }
        }
    });
    charts.push(chart);
}

function renderPromoChart(data) {
    const ctx = document.getElementById('promoChart').getContext('2d');
    const chart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(data),
            datasets: [{
                data: Object.values(data),
                backgroundColor: ['#8b5cf6', '#3b82f6', '#10b981', '#f59e0b', '#ef4444'],
                borderWidth: 0
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, cutout: '70%', plugins: { legend: { position: 'right' } } }
    });
    charts.push(chart);
}
