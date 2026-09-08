require("dotenv").config();
const express = require("express");
const cors = require("cors");
const fetch = require("node-fetch");

const app = express();

const PORT = process.env.PORT || 3000;
const BASE_API_URL = process.env.BASE_API_URL || "https://botai.smartdataautomation.com/api_backend_ai";
const REPORT_ID = process.env.REPORT_ID || "119";

// Acepta la variable con o sin el prefijo "Token " y sin espacios de sobra,
// para evitar duplicarlo al armar el header de Authorization.
function normalizeToken(raw) {
    if (!raw) return "";
    return raw.trim().replace(/^Token\s+/i, "");
}

const API_TOKEN = normalizeToken(process.env.API_TOKEN);
const JSON_TO_EXCEL_TOKEN = normalizeToken(process.env.JSON_TO_EXCEL_TOKEN);

if (!API_TOKEN) {
    console.warn("⚠️  API_TOKEN no está configurado. Define la variable de entorno API_TOKEN.");
}
if (!JSON_TO_EXCEL_TOKEN) {
    console.warn("⚠️  JSON_TO_EXCEL_TOKEN no está configurado. Define la variable de entorno JSON_TO_EXCEL_TOKEN.");
}

const REPORT_HEADERS = {
    Authorization: `Token ${API_TOKEN}`,
    "Content-Type": "application/json",
};

const JSON_TO_EXCEL_HEADERS = {
    Authorization: `Token ${JSON_TO_EXCEL_TOKEN}`,
    "Content-Type": "application/json",
};

app.use(cors());
app.use(express.json());

app.get("/health", (req, res) => {
    res.json({ status: "ok" });
});

// GET /api/export/:report?format=xlsx
// Consulta el reporte dinámico y lo convierte a excel, devolviendo el link de descarga.
app.get("/api/export/:report", async (req, res) => {
    const { report } = req.params;
    const format = req.query.format || "xlsx";

    try {
        const reportResponse = await fetch(`${BASE_API_URL}/dinamic-db/report/${REPORT_ID}/${report}`, {
            headers: REPORT_HEADERS,
        });

        if (!reportResponse.ok) {
            const text = await reportResponse.text();
            console.error(`❌ Error consultando reporte "${report}" (${reportResponse.status}):`, text);
            return res.status(reportResponse.status).json({
                error: `Error ${reportResponse.status} al consultar el reporte "${report}"`,
                details: text.substring(0, 200),
            });
        }

        const reportBody = await reportResponse.json();
        const data = Array.isArray(reportBody) ? reportBody : reportBody.result;

        if (!Array.isArray(data) || data.length === 0) {
            return res.status(404).json({ error: `El reporte "${report}" no devolvió datos` });
        }

        const excelResponse = await fetch(`${BASE_API_URL}/json_to_excel`, {
            method: "POST",
            headers: JSON_TO_EXCEL_HEADERS,
            body: JSON.stringify({ json: data, format_file: format }),
        });

        if (!excelResponse.ok) {
            const text = await excelResponse.text();
            console.error(`❌ Error generando excel para "${report}" (${excelResponse.status}):`, text);
            return res.status(excelResponse.status).json({
                error: `Error ${excelResponse.status} al generar el archivo`,
                details: text.substring(0, 200),
            });
        }

        const excelData = await excelResponse.json();
        res.json({ status: "ok", report, rows: data.length, file: excelData.file });
    } catch (err) {
        console.error(`❌ Error procesando el reporte "${report}":`, err);
        res.status(500).json({ error: "Error al procesar el reporte", details: err.message });
    }
});

app.listen(PORT, () => {
    console.log(`🚀 Servidor proxy escuchando en http://localhost:${PORT}`);
});
