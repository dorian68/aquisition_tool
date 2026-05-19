# OptiQuant IA Lead Tool Backend

FastAPI backend for the CSV to Premium Excel Dashboard Generator.

The backend owns the technical value: CSV validation, profiling, dashboard spec,
preview payload and XLSX generation. Lovable/React renders the visual preview.
n8n receives business events and never blocks generation.

## Local run

```bash
cd backend
cp .env.example .env
docker compose up --build
```

API:

- `GET /health`
- OpenAPI docs: `http://localhost:8000/docs`

For local testing, `ALLOW_MOCK_GOOGLE_AUTH=true` lets you authenticate with:

```json
{"id_token":"mock:user@example.com|User Name"}
```

## Main flow

1. `POST /api/v1/uploads/csv` with multipart `file`.
2. `POST /api/v1/uploads/{upload_id}/analyze`.
3. `POST /api/v1/uploads/{upload_id}/dashboard-spec`.
4. `GET /api/v1/dashboards/{dashboard_id}/preview`.
5. `POST /api/v1/auth/google`.
6. `POST /api/v1/dashboards/{dashboard_id}/generate-xlsx` with `Authorization: Bearer <token>`.
7. `GET /api/v1/files/{file_id}/download` with the same bearer token.

## Direct SaaS generator

For a SaaS backend that already owns upload/auth, use the direct multipart endpoint:

```bash
curl -X POST http://localhost:8000/api/v1/generator/dashboard \
  -F "file=@fixtures/sample_sales.csv;type=text/csv" \
  -F "template=auto" \
  -F "output_format=xlsx" \
  -o dashboard.xlsx
```

Supported templates:

- `auto`
- `dark-saas`
- `fintech-executive`
- `light-consulting`

Macro-enabled output:

```bash
curl -X POST http://localhost:8000/api/v1/generator/dashboard \
  -F "file=@fixtures/sample_sales.csv;type=text/csv" \
  -F "template=dark-saas" \
  -F "output_format=xlsm" \
  -F "vba_project_file=@../dashboard_generator/assets/vbaProject.bin;type=application/octet-stream" \
  -o dashboard.xlsm
```

The response body is the generated Excel file. Metadata is also returned in headers:

- `X-Dashboard-Template`
- `X-Dashboard-Dataset-Type`
- `X-Dashboard-Metadata`

## n8n

Set `N8N_ENABLED=true` and provide the relevant webhook URLs:

- `N8N_WEBHOOK_LEAD_CREATED`
- `N8N_WEBHOOK_DASHBOARD_GENERATED`
- `N8N_WEBHOOK_FILE_DOWNLOADED`

Webhook failures are logged to `n8n_webhook_logs` and do not block users.

## Tests

```bash
cd backend
pytest
```

## Fixture

Use `fixtures/sample_sales.csv` to exercise the full flow.
