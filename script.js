// Endpoint local del proxy (rama Proxy corriendo en tu máquina):
const PROXY_URL = 'http://localhost:3000';

// Endpoint de Railway (descomentar y reemplazar cuando el proxy esté desplegado):
// const PROXY_URL = 'https://TU-PROXY.up.railway.app';

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
    setLoading(true);
    renderLoading(report);

    try {
        const response = await fetch(`${PROXY_URL}/api/export/${encodeURIComponent(report)}`);
        const data = await response.json();

        if (!response.ok) {
            renderError(data.error || 'No se pudo generar el archivo.');
            return;
        }

        renderSuccess(data);
    } catch (error) {
        console.error('Error al generar el excel:', error);
        renderError('No se pudo conectar con el servidor. Verifica que el proxy esté activo.');
    } finally {
        setLoading(false);
    }
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
