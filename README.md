# Automatizacion_Data_excel

## Rama Proxy (backend)

Servidor Express que expone un endpoint para consultar un reporte dinámico y
convertirlo automáticamente a Excel, devolviendo el link de descarga.

### Endpoint

`GET /api/export/:report`

- `:report` — nombre de la base a consultar (ej. `PDVCarnicos`).
- `?format=xlsx` — opcional, formato del archivo (por defecto `xlsx`).

Respuesta:

```json
{
  "status": "ok",
  "report": "PDVCarnicos",
  "rows": 1234,
  "file": "https://s3.amazonaws.../archivo.xlsx"
}
```

Internamente hace `GET` a `dinamic-db/report/{REPORT_ID}/{report}`, pasa el
JSON resultante a `POST json_to_excel` y devuelve el link que esa API genera.

### Variables de entorno

Copiar `.env.example` a `.env` y completar:

- `PORT` — puerto local (Railway lo asigna automáticamente).
- `BASE_API_URL` — base de la API (`https://botai.smartdataautomation.com/api_backend_ai`).
- `REPORT_ID` — id del reporte dinámico (actualmente `119`).
- `API_TOKEN` — token de autorización para el GET del reporte dinámico (sin el prefijo `Token`).
- `JSON_TO_EXCEL_TOKEN` — token de autorización para el POST a `json_to_excel` (sin el prefijo `Token`, es distinto al anterior).

### Desarrollo local

```bash
npm install
npm start
```
