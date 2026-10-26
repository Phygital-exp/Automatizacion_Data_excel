// Endpoint de Railway (proxy en producción):
const PROXY_URL = 'https://automatizaciondataexcel-production.up.railway.app';

// Endpoint local del proxy (para probar en tu máquina, rama Proxy):
// const PROXY_URL = 'http://localhost:3000';

const form = document.getElementById('exportForm');
const input = document.getElementById('reportInput');
const submitBtn = document.getElementById('submitBtn');
const results = document.getElementById('results');
const panelFiltros = document.getElementById('filtros');

// Base cargada: libro original (para copiar celdas con su formato) y textos de cada celda (para filtrar)
let base = null;

const OPERADORES = [
    ['igual', 'es igual a'],
    ['contiene', 'contiene'],
    ['distinto', 'es distinto de'],
    ['no_contiene', 'no contiene'],
    ['mayor_igual', 'es mayor o igual a'],
    ['menor_igual', 'es menor o igual a'],
    ['vacio', 'está vacío'],
    ['no_vacio', 'no está vacío'],
];
const SIN_VALOR = new Set(['vacio', 'no_vacio']);

form.addEventListener('submit', (event) => {
    event.preventDefault();
    const report = input.value.trim();
    if (report) {
        buscarBase(report);
    }
});

/** Pide el reporte al proxy, lo lee en el navegador y abre el panel de filtros. */
async function buscarBase(report) {
    setLoading(true);
    renderLoading(report);
    panelFiltros.classList.add('hidden');
    base = null;

    try {
        const response = await fetch(`${PROXY_URL}/api/export/${encodeURIComponent(report)}`);
        const data = await response.json();

        if (!response.ok) {
            renderError(data.error || 'No se pudo generar el archivo.');
            return;
        }

        try {
            const archivo = await fetch(data.file);
            if (!archivo.ok) throw new Error(`HTTP ${archivo.status}`);
            const libro = XLSX.read(await archivo.arrayBuffer(), { type: 'array' });
            cargarBase(report, data, libro);
            results.innerHTML = '';
        } catch (error) {
            // Si el navegador no puede leer el archivo, queda la descarga completa de siempre
            console.error('No se pudo leer el archivo en el navegador:', error);
            renderSuccess(data);
        }
    } catch (error) {
        console.error('Error al generar el excel:', error);
        renderError('No se pudo conectar con el servidor. Verifica que el proxy esté activo.');
    } finally {
        setLoading(false);
    }
}

function cargarBase(report, data, libro) {
    const nombreHoja = libro.SheetNames[0];
    const hoja = libro.Sheets[nombreHoja];
    const ref = XLSX.utils.decode_range(hoja['!ref']);
    // raw:false da el texto tal como se ve en Excel (fechas y números formateados)
    const tabla = XLSX.utils.sheet_to_json(hoja, { header: 1, raw: false, defval: '' });
    const columnas = (tabla[0] || []).map((c, i) => String(c || `Columna ${i + 1}`));
    base = { report, data, libro, nombreHoja, hoja, ref, columnas, filas: tabla.slice(1) };

    document.getElementById('infoBase').textContent =
        `${report} · ${base.filas.length.toLocaleString('es-CO')} filas · ${columnas.length} columnas`;
    const completa = document.getElementById('descargaCompleta');
    completa.href = data.file;
    document.getElementById('listaFiltros').innerHTML = '';
    document.getElementById('filaDesde').value = '';
    document.getElementById('filaHasta').value = '';
    agregarFiltro();
    panelFiltros.classList.remove('hidden');
    actualizar();
}

/* ---------- Filtros ---------- */

let contadorFiltros = 0;

function agregarFiltro() {
    if (!base) return;
    const id = ++contadorFiltros;
    const fila = document.createElement('div');
    fila.className = 'filtro';
    fila.dataset.id = id;
    fila.innerHTML = `
        <select class="f-columna" aria-label="Columna">
            <option value="">Elige una columna…</option>
            ${base.columnas.map((c, i) => `<option value="${i}">${escapeHtml(c)}</option>`).join('')}
        </select>
        <select class="f-operador" aria-label="Condición">
            ${OPERADORES.map(([v, t]) => `<option value="${v}">${t}</option>`).join('')}
        </select>
        <input class="f-valor" type="text" list="valores-${id}" placeholder="Valor" aria-label="Valor" />
        <datalist id="valores-${id}"></datalist>
        <button type="button" class="f-quitar" aria-label="Quitar filtro" title="Quitar filtro">
            <i class="material-icons">close</i>
        </button>
    `;
    document.getElementById('listaFiltros').appendChild(fila);
}

/** Sugerencias de valores de la columna elegida, de los más frecuentes a los menos. */
function llenarSugerencias(fila) {
    const col = fila.querySelector('.f-columna').value;
    const lista = fila.querySelector('datalist');
    if (col === '') { lista.innerHTML = ''; return; }
    const cuenta = new Map();
    for (const f of base.filas) {
        const v = String(f[col] ?? '').trim();
        if (v) cuenta.set(v, (cuenta.get(v) || 0) + 1);
    }
    lista.innerHTML = [...cuenta.entries()].sort((a, b) => b[1] - a[1]).slice(0, 300)
        .map(([v, n]) => `<option value="${escapeHtml(v)}">${n.toLocaleString('es-CO')} filas</option>`).join('');
}

const normalizar = (texto) => String(texto ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();

/** Compara como número si ambos lo son; si no, como texto (sirve para fechas AAAA-MM-DD). */
function comparar(a, b) {
    const na = Number(String(a).replace(',', '.')), nb = Number(String(b).replace(',', '.'));
    if (String(a).trim() !== '' && String(b).trim() !== '' && !Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
    return normalizar(a).localeCompare(normalizar(b));
}

function cumple(valorCelda, operador, buscado) {
    const celda = normalizar(valorCelda), valor = normalizar(buscado);
    switch (operador) {
        case 'igual': return celda === valor;
        case 'distinto': return celda !== valor;
        case 'contiene': return celda.includes(valor);
        case 'no_contiene': return !celda.includes(valor);
        case 'mayor_igual': return celda !== '' && comparar(valorCelda, buscado) >= 0;
        case 'menor_igual': return celda !== '' && comparar(valorCelda, buscado) <= 0;
        case 'vacio': return celda === '';
        case 'no_vacio': return celda !== '';
        default: return true;
    }
}

/** Filtros completos (con columna y, si la condición lo pide, con valor). */
function filtrosActivos() {
    return [...document.querySelectorAll('#listaFiltros .filtro')].map((f) => ({
        columna: f.querySelector('.f-columna').value,
        operador: f.querySelector('.f-operador').value,
        valor: f.querySelector('.f-valor').value,
    })).filter((f) => f.columna !== '' && (SIN_VALOR.has(f.operador) || f.valor.trim() !== ''));
}

/** Rango de filas elegido (1 = primer registro). Devuelve {error} si no es válido. */
function leerRangoFilas() {
    const textoDesde = document.getElementById('filaDesde').value.trim();
    const textoHasta = document.getElementById('filaHasta').value.trim();
    const desde = textoDesde ? Number(textoDesde) : 1;
    const hasta = textoHasta ? Number(textoHasta) : Infinity;
    if (!Number.isInteger(desde) || desde < 1 || (hasta !== Infinity && (!Number.isInteger(hasta) || hasta < 1))) {
        return { error: 'Las filas deben ser números enteros mayores o iguales a 1.' };
    }
    if (desde > hasta) return { error: '"Desde la fila" no puede ser mayor que "Hasta la fila".' };
    return { activo: Boolean(textoDesde || textoHasta), desde, hasta };
}

/** Índices (desde 0) de las filas de datos que pasan el rango y todos los filtros. */
function filasFiltradas() {
    const rango = leerRangoFilas();
    if (rango.error) return { error: rango.error };
    const filtros = filtrosActivos();
    const indices = [];
    const fin = Math.min(rango.hasta, base.filas.length);
    for (let i = rango.desde - 1; i < fin; i++) {
        const fila = base.filas[i];
        if (filtros.every((f) => cumple(fila[f.columna], f.operador, f.valor))) indices.push(i);
    }
    return { indices, filtros, rango };
}

/** Recalcula el conteo y la vista previa cada vez que cambia un filtro. */
function actualizar() {
    if (!base) return;
    document.querySelectorAll('#listaFiltros .filtro').forEach((f) => {
        f.querySelector('.f-valor').classList.toggle('hidden', SIN_VALOR.has(f.querySelector('.f-operador').value));
    });
    const conteo = document.getElementById('conteo');
    const r = filasFiltradas();
    const boton = document.getElementById('descargarFiltrado');
    if (r.error) {
        conteo.innerHTML = `<span class="error-texto">${escapeHtml(r.error)}</span>`;
        boton.disabled = true;
        document.getElementById('vista').innerHTML = '';
        return;
    }
    boton.disabled = r.indices.length === 0;
    conteo.innerHTML = `<b>${r.indices.length.toLocaleString('es-CO')}</b> de ${base.filas.length.toLocaleString('es-CO')} filas cumplen` +
        (r.filtros.length || r.rango.activo ? '' : ' (sin filtros: se descarga todo)');
    const muestra = r.indices.slice(0, 5);
    document.getElementById('vista').innerHTML = muestra.length
        ? `<thead><tr><th>Fila</th>${base.columnas.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}</tr></thead><tbody>` +
          muestra.map((i) => `<tr><td>${i + 1}</td>${base.columnas.map((_, c) => `<td>${escapeHtml(base.filas[i][c] ?? '')}</td>`).join('')}</tr>`).join('') +
          '</tbody>'
        : '';
}

document.getElementById('agregarFiltro').addEventListener('click', () => { agregarFiltro(); actualizar(); });
document.getElementById('listaFiltros').addEventListener('change', (e) => {
    if (e.target.classList.contains('f-columna')) llenarSugerencias(e.target.closest('.filtro'));
    actualizar();
});
document.getElementById('listaFiltros').addEventListener('input', actualizar);
document.getElementById('listaFiltros').addEventListener('click', (e) => {
    const quitar = e.target.closest('.f-quitar');
    if (!quitar) return;
    quitar.closest('.filtro').remove();
    if (!document.querySelector('#listaFiltros .filtro')) agregarFiltro();
    actualizar();
});
['filaDesde', 'filaHasta'].forEach((id) => document.getElementById(id).addEventListener('input', actualizar));

/* ---------- Descarga ---------- */

const limpiarNombre = (texto) => String(texto).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^A-Za-z0-9-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40);

/** Nombre del archivo: reporte + valores filtrados (con «es igual a») + rango de filas. */
function nombreArchivo(filtros, rango, ultima) {
    const partes = [limpiarNombre(base.report)];
    const valores = filtros.filter((f) => f.operador === 'igual').map((f) => limpiarNombre(f.valor)).filter(Boolean);
    if (valores.length) partes.push(valores.join('_'));
    else if (filtros.length) partes.push('filtrado');
    if (rango.activo) partes.push(`${rango.desde}-${ultima}`);
    return `${partes.join('_').slice(0, 120)}.xlsx`;
}

/** Copia los títulos y las filas elegidas celda por celda (tipos y formatos intactos) y descarga. */
document.getElementById('descargarFiltrado').addEventListener('click', () => {
    const r = filasFiltradas();
    if (r.error || !r.indices.length) return;
    const { hoja, ref } = base;
    const nueva = {};
    const copiarFila = (origen, destino) => {
        for (let c = ref.s.c; c <= ref.e.c; c++) {
            const celda = hoja[XLSX.utils.encode_cell({ r: origen, c })];
            if (celda) nueva[XLSX.utils.encode_cell({ r: destino, c })] = { ...celda };
        }
    };
    copiarFila(ref.s.r, 0);
    r.indices.forEach((i, k) => copiarFila(ref.s.r + 1 + i, k + 1));
    nueva['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: ref.s.c }, e: { r: r.indices.length, c: ref.e.c } });
    if (hoja['!cols']) nueva['!cols'] = hoja['!cols'];
    const salida = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(salida, nueva, base.nombreHoja.slice(0, 31));
    const ultima = Math.min(r.rango.hasta, base.filas.length);
    const nombre = nombreArchivo(r.filtros, r.rango, ultima);
    XLSX.writeFile(salida, nombre);
    results.innerHTML = `
        <div class="state-card success">
            <i class="material-icons">check_circle</i>
            <div class="body">
                <div class="title">Archivo descargado: ${escapeHtml(nombre)}</div>
                <div>${r.indices.length.toLocaleString('es-CO')} de ${base.filas.length.toLocaleString('es-CO')} filas de ${escapeHtml(base.report)}</div>
            </div>
        </div>
    `;
});

/* ---------- Estados de la página ---------- */

function setLoading(isLoading) {
    submitBtn.disabled = isLoading;
    input.disabled = isLoading;
}

function renderLoading(report) {
    results.innerHTML = `
        <div class="state-card loading">
            <div class="spinner"></div>
            <div class="body">
                <div class="title">Buscando "${escapeHtml(report)}"...</div>
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
