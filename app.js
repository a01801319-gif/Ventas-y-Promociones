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

// SKU Selector Elements
const skuSelect = document.getElementById('sku-select');

// Datos Globales y Estado
let dataVentas = null;
let dataPromos = null;
let charts = [];
let skuChartInstance = null; // Para guardar la gráfica de SKU separadamente

let state = {
    groupBy: 'day', // 'day' | 'month'
    chartType: 'line' // 'line' | 'bar'
};

// Almacenamos los datos procesados limpios para usarlos en filtros rápidos
let globalSalesData = [];

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
        processData(false);
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

skuSelect.addEventListener('change', (e) => {
    renderSkuDeepDiveChart(e.target.value);
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
    globalSalesData = [];
    statusVentas.textContent = "Esperando archivo..."; statusVentas.className = "file-status";
    statusPromos.textContent = "Esperando archivo..."; statusPromos.className = "file-status";
    btnAnalyze.disabled = true;
    charts.forEach(c => c.destroy());
    charts = [];
    if(skuChartInstance) skuChartInstance.destroy();
    aiContainer.innerHTML = '<p class="ai-placeholder">Esperando a la IA...</p>';
});

// --- LÓGICA PRINCIPAL ---
function processData(initialLoad = false) {
    try {
        if (!dataVentas || dataVentas.length === 0) throw new Error("Base vacía");
        
        const vRow = dataVentas[0];
        const colFecha = findColumn(vRow, ['fecha', 'date', 'dia']);
        const colProducto = findColumn(vRow, ['producto', 'product', 'item', 'sku']);
        const colVenta = findColumn(vRow, ['venta', 'sales', 'monto', 'ingreso', 'revenue']);
        
        // Nuevas columnas para el Deep Dive
        const colUnidades = findColumn(vRow, ['unidad', 'unit', 'qty', 'cantidad']);
        const colGanancia = findColumn(vRow, ['ganancia', 'profit', 'margen', 'utilidad', 'costo']); // Costo se podría restar, pero asumimos ganancia directa
        
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
        const timelineData = {};
        const productData = {};
        const promoPerformance = {};
        const uniqueSkus = new Set();
        
        if(initialLoad) globalSalesData = []; // Resetear solo en la carga inicial

        dataVentas.forEach(v => {
            const val = parseFloat(v[colVenta]) || 0;
            const prod = String(v[colProducto] || 'Desconocido').trim();
            const rawDate = String(v[colFecha] || 'Sin Fecha').trim();
            
            const unidades = parseFloat(v[colUnidades]) || 1; // Fallback a 1
            const ganancia = parseFloat(v[colGanancia]) || (val * 0.3); // Fallback a 30% margen si no existe
            
            uniqueSkus.add(prod);

            if(initialLoad) {
                globalSalesData.push({
                    date: rawDate,
                    product: prod,
                    revenue: val,
                    units: unidades,
                    profit: ganancia
                });
            }

            // Agrupación de fechas
            let dateKey = rawDate;
            if (state.groupBy === 'month' && rawDate.length >= 7) {
                const match = rawDate.match(/^(\d{4}[-/]\d{2})/);
                if (match) dateKey = match[1].replace('/', '-');
                else dateKey = rawDate;
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

        const chartTimelineData = {};
        Object.keys(timelineData).forEach(key => {
            if (state.groupBy === 'month') {
                const count = timelineData[key].count || 1;
                chartTimelineData[key] = {
                    promo: timelineData[key].promo / count,
                    normal: timelineData[key].normal / count
                };
            } else {
                chartTimelineData[key] = timelineData[key];
            }
        });

        // Actualizar KPIs
        document.getElementById('kpi-total-sales').textContent = `$${totalSales.toLocaleString('es-MX', {maximumFractionDigits:0})}`;
        document.getElementById('kpi-promo-sales').textContent = `$${promoSales.toLocaleString('es-MX', {maximumFractionDigits:0})}`;
        const promoPct = totalSales > 0 ? ((promoSales / totalSales) * 100).toFixed(1) : 0;
        document.getElementById('kpi-promo-percentage').textContent = `${promoPct}% del total`;
        const normalSales = totalSales - promoSales;
        const lift = normalSales > 0 ? ((promoSales / normalSales) * 100).toFixed(1) : 0;
        document.getElementById('kpi-lift').textContent = `+${lift}%`;

        // Llenar Selector de SKU solo en la carga inicial
        if (initialLoad) {
            skuSelect.innerHTML = '<option value="">-- Selecciona un producto --</option>';
            Array.from(uniqueSkus).sort().forEach(sku => {
                const opt = document.createElement('option');
                opt.value = sku;
                opt.textContent = sku;
                skuSelect.appendChild(opt);
            });
        }

        // Destruir gráficas viejas
        charts.forEach(c => c.destroy());
        charts = [];

        // Render
        renderTimelineChart(chartTimelineData);
        renderProductChart(productData);
        renderPromoChart(promoPerformance);
        
        // Auto-renderizar Deep Dive si hay uno seleccionado
        if (skuSelect.value) {
            renderSkuDeepDiveChart(skuSelect.value);
        } else if (initialLoad && uniqueSkus.size > 0) {
            // Seleccionar el primero por defecto
            const firstSku = Array.from(uniqueSkus).sort()[0];
            skuSelect.value = firstSku;
            renderSkuDeepDiveChart(firstSku);
        }

        // Llamar a la IA
        if (initialLoad) {
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

// --- GRÁFICA DE DEEP DIVE (MULTI-EJE) ---
function renderSkuDeepDiveChart(sku) {
    if (!sku) return;
    
    // Filtrar datos para este SKU
    const skuData = globalSalesData.filter(item => item.product === sku);
    
    // Agrupar por fecha cronológicamente
    const aggregated = {};
    skuData.forEach(item => {
        // En Deep Dive siempre mostramos días, o respetamos el groupBy?
        // Respetemos el groupBy para consistencia
        let dateKey = item.date;
        if (state.groupBy === 'month' && item.date.length >= 7) {
            const match = item.date.match(/^(\d{4}[-/]\d{2})/);
            if (match) dateKey = match[1].replace('/', '-');
        }

        if (!aggregated[dateKey]) {
            aggregated[dateKey] = { revenue: 0, units: 0, profit: 0 };
        }
        aggregated[dateKey].revenue += item.revenue;
        aggregated[dateKey].units += item.units;
        aggregated[dateKey].profit += item.profit;
    });

    const labels = Object.keys(aggregated).sort();
    const revenueData = labels.map(l => aggregated[l].revenue);
    const unitsData = labels.map(l => aggregated[l].units);
    const profitData = labels.map(l => aggregated[l].profit);

    if (skuChartInstance) {
        skuChartInstance.destroy();
    }

    const ctx = document.getElementById('skuDeepDiveChart').getContext('2d');
    skuChartInstance = new Chart(ctx, {
        type: 'bar', // Base chart type
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Unidades Vendidas',
                    data: unitsData,
                    type: 'bar',
                    backgroundColor: 'rgba(59, 130, 246, 0.6)', // Azul secundario
                    yAxisID: 'y1',
                    order: 3
                },
                {
                    label: 'Ingresos ($)',
                    data: revenueData,
                    type: 'line',
                    borderColor: '#8b5cf6', // Morado primario
                    backgroundColor: '#8b5cf6',
                    borderWidth: 3,
                    tension: 0.3,
                    yAxisID: 'y',
                    order: 2
                },
                {
                    label: 'Ganancia ($)',
                    data: profitData,
                    type: 'line',
                    borderColor: '#10b981', // Verde success
                    backgroundColor: '#10b981',
                    borderDash: [5, 5],
                    borderWidth: 3,
                    tension: 0.3,
                    yAxisID: 'y',
                    order: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            scales: {
                x: {
                    grid: { display: false }
                },
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    title: { display: true, text: 'Dinero ($)', color: '#94a3b8' },
                    grid: { color: 'rgba(255,255,255,0.05)' }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    title: { display: true, text: 'Unidades (Cant.)', color: '#94a3b8' },
                    grid: { drawOnChartArea: false } // no dibujar lineas para no empalmar
                }
            }
        }
    });
}

// --- GRÁFICAS EXISTENTES ---
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
