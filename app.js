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

// Datos Globales
let dataVentas = null;
let dataPromos = null;
let charts = []; // Para poder destruirlas al resetear

// Helpers para Drag & Drop
function setupDropzone(dropzone, input, statusElem, type) {
    dropzone.addEventListener('click', () => input.click());
    
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
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0], statusElem, type);
        }
    });

    input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0], statusElem, type);
        }
    });
}

setupDropzone(dropzoneVentas, fileVentas, statusVentas, 'ventas');
setupDropzone(dropzonePromos, filePromos, statusPromos, 'promos');

// Procesar Archivo con SheetJS
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
            
            if (type === 'ventas') {
                dataVentas = json;
            } else {
                dataPromos = json;
            }
            
            statusElem.textContent = `✅ ${file.name} cargado`;
            statusElem.className = "file-status loaded";
            
            checkReady();
        } catch (error) {
            statusElem.textContent = "❌ Error al leer el archivo";
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function checkReady() {
    if (dataVentas && dataPromos) {
        btnAnalyze.disabled = false;
    }
}

// Búsqueda difusa de nombres de columnas
function findColumn(row, keywords) {
    const keys = Object.keys(row);
    for (let key of keys) {
        const lowerKey = key.toLowerCase();
        if (keywords.some(kw => lowerKey.includes(kw))) {
            return key;
        }
    }
    return null;
}

// Lógica Principal de Análisis
btnAnalyze.addEventListener('click', () => {
    uploadSection.classList.add('hidden');
    loader.classList.remove('hidden');
    
    // Simular un pequeño tiempo de carga para el efecto "Wow"
    setTimeout(() => {
        processAndRenderDashboard();
    }, 1000);
});

btnReset.addEventListener('click', () => {
    dashboardSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    dataVentas = null;
    dataPromos = null;
    statusVentas.textContent = "Esperando archivo...";
    statusVentas.className = "file-status";
    statusPromos.textContent = "Esperando archivo...";
    statusPromos.className = "file-status";
    btnAnalyze.disabled = true;
    charts.forEach(c => c.destroy());
    charts = [];
});

function processAndRenderDashboard() {
    try {
        if (!dataVentas || dataVentas.length === 0) throw new Error("Base de ventas vacía");
        
        // Identificar columnas clave en Ventas
        const vRow = dataVentas[0];
        const colFecha = findColumn(vRow, ['fecha', 'date', 'dia']);
        const colProducto = findColumn(vRow, ['producto', 'product', 'item', 'sku']);
        const colVenta = findColumn(vRow, ['venta', 'sales', 'monto', 'ingreso', 'revenue', 'cantidad']);
        
        // Identificar columnas en Promos
        const pRow = dataPromos[0] || {};
        const colPromoName = findColumn(pRow, ['promo', 'campaña', 'campaign', 'nombre']);
        const colPromoProd = findColumn(pRow, ['producto', 'product', 'item', 'sku']);
        
        // Vamos a cruzar las tablas de forma simplificada:
        // Asumimos que dataPromos tiene [Producto, Promocion, FechaInicio (opcional)]
        // Si no hay fechas, asumimos que aplicó siempre. Para simplificar el MVP, 
        // mapeamos Productos a sus Promociones.
        
        const promoMap = {};
        dataPromos.forEach(p => {
            const prod = String(p[colPromoProd || colProducto] || '').trim().toLowerCase();
            const promo = p[colPromoName] || 'Sí';
            if (prod) promoMap[prod] = promo;
        });

        // Procesar Ventas
        let totalSales = 0;
        let promoSales = 0;
        const timelineData = {};
        const productData = {};
        const promoPerformance = {};

        dataVentas.forEach(v => {
            const val = parseFloat(v[colVenta]) || 0;
            const prod = String(v[colProducto] || 'Desconocido').trim();
            const date = v[colFecha] || 'Sin Fecha';
            const prodLower = prod.toLowerCase();
            
            const hasPromo = promoMap[prodLower];
            const promoName = hasPromo ? promoMap[prodLower] : 'Sin Promoción';

            totalSales += val;
            if (hasPromo) promoSales += val;

            // Timeline
            if (!timelineData[date]) timelineData[date] = { normal: 0, promo: 0 };
            if (hasPromo) timelineData[date].promo += val;
            else timelineData[date].normal += val;

            // Productos
            if (!productData[prod]) productData[prod] = { normal: 0, promo: 0 };
            if (hasPromo) productData[prod].promo += val;
            else productData[prod].normal += val;

            // Rendimiento de Promos
            if (hasPromo) {
                if (!promoPerformance[promoName]) promoPerformance[promoName] = 0;
                promoPerformance[promoName] += val;
            }
        });

        // Actualizar KPIs
        document.getElementById('kpi-total-sales').textContent = `$${totalSales.toLocaleString('es-MX', {maximumFractionDigits:0})}`;
        document.getElementById('kpi-promo-sales').textContent = `$${promoSales.toLocaleString('es-MX', {maximumFractionDigits:0})}`;
        
        const promoPct = totalSales > 0 ? ((promoSales / totalSales) * 100).toFixed(1) : 0;
        document.getElementById('kpi-promo-percentage').textContent = `${promoPct}% de los ingresos totales`;

        // Calcular Lift aproximado (ventas de productos con promo vs sin promo)
        // Esto es un insight básico.
        const normalSales = totalSales - promoSales;
        const lift = normalSales > 0 ? ((promoSales / normalSales) * 100).toFixed(1) : 0;
        document.getElementById('kpi-lift').textContent = `+${lift}%`;

        // Renderizar Gráficas
        renderTimelineChart(timelineData);
        renderProductChart(productData);
        renderPromoChart(promoPerformance);

        // Generar Insights
        generateInsights(promoSales, normalSales, promoPerformance, productData, totalSales);

        loader.classList.add('hidden');
        dashboardSection.classList.remove('hidden');

    } catch (err) {
        alert("Ocurrió un error procesando los datos. Asegúrate de que los archivos tengan columnas de Producto y Venta.\nDetalle: " + err.message);
        loader.classList.add('hidden');
        uploadSection.classList.remove('hidden');
    }
}

// Configuración Global de Chart.js
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = 'Outfit';

function renderTimelineChart(data) {
    const ctx = document.getElementById('timelineChart').getContext('2d');
    const labels = Object.keys(data).sort(); // Ordenar fechas
    const promoData = labels.map(l => data[l].promo);
    const normalData = labels.map(l => data[l].normal);

    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Ventas en Promoción',
                    data: promoData,
                    borderColor: '#8b5cf6',
                    backgroundColor: 'rgba(139, 92, 246, 0.2)',
                    fill: true,
                    tension: 0.4
                },
                {
                    label: 'Ventas Normales',
                    data: normalData,
                    borderColor: '#3b82f6',
                    backgroundColor: 'transparent',
                    tension: 0.4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { position: 'top' } },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
                x: { grid: { display: false } }
            }
        }
    });
    charts.push(chart);
}

function renderProductChart(data) {
    const ctx = document.getElementById('productChart').getContext('2d');
    // Top 10 productos
    const sortedProds = Object.keys(data).sort((a,b) => (data[b].promo + data[b].normal) - (data[a].promo + data[a].normal)).slice(0, 7);
    
    const promoData = sortedProds.map(p => data[p].promo);
    const normalData = sortedProds.map(p => data[p].normal);

    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sortedProds,
            datasets: [
                {
                    label: 'Con Promo',
                    data: promoData,
                    backgroundColor: '#8b5cf6',
                    borderRadius: 4
                },
                {
                    label: 'Sin Promo',
                    data: normalData,
                    backgroundColor: '#3b82f6',
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true, grid: { display: false } },
                y: { stacked: true, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
        }
    });
    charts.push(chart);
}

function renderPromoChart(data) {
    const ctx = document.getElementById('promoChart').getContext('2d');
    const labels = Object.keys(data);
    const values = Object.values(data);

    const chart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: values,
                backgroundColor: [
                    '#8b5cf6', '#3b82f6', '#10b981', '#f59e0b', '#ef4444'
                ],
                borderWidth: 0,
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right' }
            },
            cutout: '70%'
        }
    });
    charts.push(chart);
}

function generateInsights(promoSales, normalSales, promoPerf, prodData, totalSales) {
    const list = document.getElementById('insights-list');
    list.innerHTML = ''; // Limpiar

    const addInsight = (text) => {
        const li = document.createElement('li');
        li.textContent = text;
        list.appendChild(li);
    };

    // Insight 1: Proporción
    const pct = totalSales > 0 ? Math.round((promoSales / totalSales) * 100) : 0;
    addInsight(`📈 El ${pct}% de los ingresos de este periodo provino de productos con promoción activa.`);

    // Insight 2: Mejor promo
    const promos = Object.keys(promoPerf);
    if (promos.length > 0) {
        const bestPromo = promos.reduce((a, b) => promoPerf[a] > promoPerf[b] ? a : b);
        addInsight(`🏆 La promoción más exitosa fue "${bestPromo}", generando $${promoPerf[bestPromo].toLocaleString('es-MX', {maximumFractionDigits:0})}.`);
    }

    // Insight 3: Dependencia
    if (promoSales > normalSales) {
        addInsight(`⚠️ Alta dependencia: Las ventas bajo promoción superaron a las ventas regulares. Sugerimos revisar márgenes de ganancia.`);
    } else {
        addInsight(`✅ Base sólida: Las ventas regulares (sin promoción) mantienen el mayor volumen de ingresos.`);
    }
}
