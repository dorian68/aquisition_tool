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

## AI Dashboard Analyst

Preview the compact Python-generated profile and AI/fallback report:

```bash
curl -X POST http://localhost:8000/api/v1/generator/analyze \
  -F "file=@fixtures/sample_sales.csv;type=text/csv" \
  -F "template=auto" \
  -F "output_format=xlsx"
```

Generate with AI analysis embedded:

```bash
curl -X POST http://localhost:8000/api/v1/generator/dashboard \
  -F "file=@fixtures/sample_sales.csv;type=text/csv" \
  -F "template=auto" \
  -F "output_format=xlsx" \
  -F "include_ai_analysis=true" \
  -o dashboard.xlsx
```

Disable AI analysis while keeping the old behavior:

```bash
curl -X POST http://localhost:8000/api/v1/generator/dashboard \
  -F "file=@fixtures/sample_sales.csv;type=text/csv" \
  -F "template=dark-saas" \
  -F "output_format=xlsx" \
  -F "include_ai_analysis=false" \
  -o dashboard.xlsx
```

The AI Analyst never receives raw CSV rows. Python sends only a capped context JSON containing profile, quality, aggregation and anomaly summaries. LangGraph orchestrates the LLM report generation inside the backend. If `OPENAI_API_KEY` is missing, AI is disabled, or the LLM fails schema validation, the dashboard still generates with deterministic fallback insights.

Relevant environment variables:

```env
AI_ANALYST_ENABLED=true
AI_ANALYST_PROVIDER=langgraph
AI_ANALYST_MODEL=gpt-4.1-mini
AI_ANALYST_TIMEOUT_SECONDS=25
MAX_UPLOAD_MB=25
MAX_COLUMNS=100
MAX_AI_CONTEXT_CHARS=20000
MAX_TOP_CATEGORIES=10
MAX_TREND_POINTS=12
MAX_ANOMALIES=5
MAX_CLEANING_ACTIONS_FOR_AI=20
MAX_HIGH_MISSING_COLUMNS_FOR_AI=10
AI_ANALYST_FAIL_OPEN=true
OPENAI_API_KEY=
```

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
