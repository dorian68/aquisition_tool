# aquisition_tool

FastAPI + Docker service that turns uploaded CSV files into premium Excel dashboards.

## Direct generator endpoint

```bash
curl -X POST http://localhost:8000/api/v1/generator/dashboard \
  -F "file=@backend/fixtures/sample_sales.csv;type=text/csv" \
  -F "template=auto" \
  -F "output_format=xlsx" \
  -o dashboard.xlsx
```

Supported templates:

- `auto`
- `dark-saas`
- `fintech-executive`
- `light-consulting`

## Docker

```bash
cd backend
docker compose up --build -d
```

API docs are available at `http://localhost:8000/docs`.
