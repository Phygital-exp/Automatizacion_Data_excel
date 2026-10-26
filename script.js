// Endpoint de Railway (proxy en producción):
const PROXY_URL = 'https://automatizaciondataexcel-production.up.railway.app';

// Endpoint local del proxy (para probar en tu máquina, rama Proxy):
// const PROXY_URL = 'http://localhost:3000';

const form = document.getElementById('exportForm');
const input = document.getElementById('reportInput');
const submitBtn = document.getElementById('submitBtn');
const results = document.getElementById('results');

form.addEventListener('submit', (event) => {
    event.preventDefault();
    const report = input.value.trim();
    if (report) {
        generarExcel(report);
    }
});

async function generarExcel(report) {
    const rango = leerRangoFilas();
    if (rango.error) {
        renderError(rango.error);
        return;
    }
    setLoading(true);
    renderLoading(report);

    try {
        const response = await fetch(`${PROXY_URL}/api/export/${encodeURIComponent(report)}`);
        const data = await response.json();

        if (!response.ok) {
            renderError(data.error || 'No se pudo generar el archivo.');
            return;
        }

        if (!rango.activo) {
            renderSuccess(data);
            return;
        }
        // Con rango de filas el archivo se lee en el navegador y se recorta; el proxy no cambia
        const archivo = await fetch(data.file);
        if (!archivo.ok) {
            renderError(`No se pudo leer el archivo generado (HTTP ${archivo.status}).`);
            return;
        }
        const libro = XLSX.read(await archivo.arrayBuffer(), { type: 'array' });
        descargarRango(report, libro, rango);
    } catch (error) {
        console.error('Error al generar el excel:', error);
        renderError('No se pudo conectar con el servidor. Verifica que el proxy esté activo.');
    } finally {
        setLoading(false);
    }
}

/* ---------- Filtro por rango de filas ---------- */

/** Lee "Desde la fila" / "Hasta la fila". La fila 1 es el primer registro debajo de los títulos. */
function leerRangoFilas() {
    const textoDesde = document.getElementById('filaDesde').value.trim();
    const textoHasta = document.getElementById('filaHasta').value.trim();
    if (!textoDesde && !textoHasta) return { activo: false };
    const desde = textoDesde ? Number(textoDesde) : 1;
    const hasta = textoHasta ? Number(textoHasta) : Infinity;
    if (!Number.isInteger(desde) || desde < 1 || (hasta !== Infinity && (!Number.isInteger(hasta) || hasta < 1))) {
        return { error: 'Las filas deben ser números enteros mayores o iguales a 1.' };
    }
    if (desde > hasta) {
        return { error: '"Desde la fila" no puede ser mayor que "Hasta la fila".' };
    }
    return { activo: true, desde, hasta };
}

/** Copia los títulos y las filas del rango (celda por celda, con tipos y formatos intactos) y descarga el archivo. */
function descargarRango(report, libro, rango) {
    const nombreHoja = libro.SheetNames[0];
    const hoja = libro.Sheets[nombreHoja];
    const ref = XLSX.utils.decode_range(hoja['!ref']);
    const total = ref.e.r - ref.s.r;   // filas de datos, sin contar los títulos
    if (rango.desde > total) {
        renderError(`El reporte "${report}" tiene ${total.toLocaleString('es-CO')} filas; "Desde la fila" (${rango.desde}) es mayor.`);
        return;
    }
    const hasta = Math.min(rango.hasta, total);
    const nueva = {};
    const copiarFila = (origen, destino) => {
        for (let c = ref.s.c; c <= ref.e.c; c++) {
            const celda = hoja[XLSX.utils.encode_cell({ r: origen, c })];
            if (celda) nueva[XLSX.utils.encode_cell({ r: destino, c })] = { ...celda };
        }
    };
    copiarFila(ref.s.r, 0);
    for (let fila = rango.desde; fila <= hasta; fila++) copiarFila(ref.s.r + fila, fila - rango.desde + 1);
    nueva['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: ref.s.c }, e: { r: hasta - rango.desde + 1, c: ref.e.c } });
    if (hoja['!cols']) nueva['!cols'] = hoja['!cols'];

    const salida = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(salida, nueva, nombreHoja.slice(0, 31));
    const nombre = `${report.replace(/[\\/:*?"<>|\s]+/g, '_')}_${rango.desde}-${hasta}.xlsx`;
    XLSX.writeFile(salida, nombre);
    renderRango({ report, nombre, desde: rango.desde, hasta, total, recortado: rango.hasta !== Infinity && rango.hasta > total, salida });
}

function renderRango(r) {
    const cantidad = r.hasta - r.desde + 1;
    results.innerHTML = `
        <div class="state-card success">
            <i class="material-icons">check_circle</i>
            <div class="body">
                <div class="title">Archivo descargado: ${escapeHtml(r.nombre)}</div>
                <div>Filas ${r.desde.toLocaleString('es-CO')} a ${r.hasta.toLocaleString('es-CO')} (${cantidad.toLocaleString('es-CO')} de ${r.total.toLocaleString('es-CO')} filas del reporte ${escapeHtml(r.report)})</div>
                ${r.recortado ? `<div class="detalle">El reporte solo tiene ${r.total.toLocaleString('es-CO')} filas; se descargó hasta la última.</div>` : ''}
                <button type="button" class="download-link" id="redescargar">
                    <i class="material-icons">file_download</i>
                    Descargar de nuevo
                </button>
            </div>
        </div>
    `;
    document.getElementById('redescargar').addEventListener('click', () => XLSX.writeFile(r.salida, r.nombre));
}

function setLoading(isLoading) {
    submitBtn.disabled = isLoading;
    input.disabled = isLoading;
}

function renderLoading(report) {
    results.innerHTML = `
        <div class="state-card loading">
            <div class="spinner"></div>
            <div class="body">
                <div class="title">Generando "${escapeHtml(report)}"...</div>
                <div>Esto puede tardar unos segundos.</div>
            </div>
        </div>
    `;
}

function renderSuccess(data) {
    results.innerHTML = `
        <div class="state-card success">
            <i class="material-icons">check_circle</i>
            <div class="body">
                <div class="title">Archivo listo (${data.rows} filas)</div>
                <div>Reporte: ${escapeHtml(data.report)}</div>
                <a class="download-link" href="${data.file}" target="_blank" rel="noopener noreferrer">
                    <i class="material-icons">file_download</i>
                    Descargar Excel
                </a>
            </div>
        </div>
    `;
}

function renderError(message) {
    results.innerHTML = `
        <div class="state-card error">
            <i class="material-icons">error_outline</i>
            <div class="body">
                <div class="title">No se pudo generar el archivo</div>
                <div>${escapeHtml(message)}</div>
            </div>
        </div>
    `;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

function toggleDarkMode() {
    const isDark = document.getElementById('darkModeToggle').checked;
    document.body.classList.toggle('dark-mode', isDark);
    document.getElementById('darkModeLabel').textContent = isDark ? '☀️ Modo claro' : '🌙 Modo oscuro';
    localStorage.setItem('darkMode', isDark ? 'enabled' : 'disabled');
}

window.onload = () => {
    const darkModeSetting = localStorage.getItem('darkMode');
    const isDark = darkModeSetting === 'enabled';
    document.body.classList.toggle('dark-mode', isDark);
    document.getElementById('darkModeToggle').checked = isDark;
    document.getElementById('darkModeLabel').textContent = isDark ? '☀️ Modo claro' : '🌙 Modo oscuro';
};
