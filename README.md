# Planner v10.0 — Universal Data Engineering Service

A scalable, multi-format task aggregation and reporting service. Replaces the legacy Tkinter desktop tool with a modern web-first architecture.

## Architecture

```
planer_v2/
├── app.py                        # Streamlit dashboard (main UI)
├── main.py                       # CLI entry point & smoke test
├── pyproject.toml                # uv dependency management
├── Dockerfile                    # Container deployment
├── src/
│   ├── engine/
│   │   ├── processor.py          # PlannerEngine — headless data processing
│   │   ├── ingestor.py           # UniversalIngestor — multi-format input
│   │   └── excel_writer.py       # ExcelReportWriter — xlsxwriter reports
│   ├── api/
│   │   └── service.py            # FastAPI REST endpoints
│   └── ai/
│       └── analyst.py            # LLM analysis hook (Ollama/LiteLLM)
├── tests/
│   └── test_engine.py            # Unit tests
├── data/                         # Working data directory
└── assets/                       # Static assets
```

## Quick Start

```bash
# Install dependencies
uv sync

# Launch Streamlit dashboard
uv run streamlit run app.py

# Or launch FastAPI service
uv sync --extra api
uv run uvicorn src.api.service:app --reload --port 8000

# Run tests
uv run pytest tests/ -v

# Smoke test
uv run python main.py
```

## Features

### Engine Layer
- **PlannerEngine**: Pure headless class — operates on DataFrames, zero UI dependencies
- **UniversalIngestor**: Supports `.xlsx`, `.xls`, `.xlsm`, `.csv`, `.json`, `.txt`
- **Hierarchical Pivot**: Osoba > Rok > Kwartał with Ukończone/Nieukończone metrics
- **Sentinel date filtering** and **unique-key deduplication** preserved from legacy

### Streamlit Dashboard
- **Ingest**: Drag & drop multiple files + raw text paste
- **Analytics**: Interactive Plotly charts with hierarchical X-axis
- **Export**: Formatted Excel (multi-level category axis) + CSV/JSON downloads
- Dark mode, responsive layout

### FastAPI Service
- `POST /ingest` — Upload & auto-detect columns
- `POST /merge` — Merge new data with dedup
- `POST /run` — Execute aggregation pipeline
- `POST /export/excel` — Download formatted report
- `POST /analyze` — AI-powered insights

### AI Analysis
- Connects to local LLM via LiteLLM (Ollama, OpenAI-compatible)
- Generates Polish-language insights and recommendations
- Graceful degradation when LLM unavailable

## Docker Deployment

```bash
docker build -t planner:v10 .
docker run -p 8501:8501 planner:v10
```

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `PLANNER_LLM_MODEL` | `ollama/llama3.1:8b` | LLM model identifier |
| `PLANNER_LLM_BASE_URL` | `http://localhost:11434` | LLM API endpoint |
| `PLANNER_LLM_API_KEY` | _(empty)_ | API key (if required) |