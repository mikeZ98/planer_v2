"""
FastAPI Service v10.0 — REST API for the Planner Engine.

Endpoints:
  POST /ingest          — Upload file(s), returns column mapping + preview
  POST /merge           — Merge new data into master
  POST /run             — Execute aggregation pipeline
  POST /export/excel    — Generate and download Excel report
  POST /analyze         — AI-powered analysis (requires LLM)
  GET  /health          — Health check
"""

from __future__ import annotations

import io
import tempfile
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import Response, StreamingResponse
from pydantic import BaseModel

from ..engine.excel_writer import ExcelReportWriter
from ..engine.ingestor import UniversalIngestor
from ..engine.processor import ColumnMap, EngineResult, PlannerEngine

app = FastAPI(
    title="Planner Engine API",
    version="10.0.0",
    description="Universal Data Engineering Service — task aggregation & reporting.",
)

# ── In-memory engine state (per-process; for production use Redis/DB) ────────
_engine: Optional[PlannerEngine] = None
_last_result: Optional[EngineResult] = None


# ── Models ───────────────────────────────────────────────────────────────────

class ColumnMapRequest(BaseModel):
    person: str
    created: str
    completed: str
    task: Optional[str] = None


class IngestResponse(BaseModel):
    rows: int
    columns: list[str]
    auto_mapping: dict[str, Optional[str]]
    preview: list[dict]


class MergeResponse(BaseModel):
    new_file_rows: int
    after_date_filter: int
    duplicates_skipped: int
    rows_added: int
    total_rows: int


class RunResponse(BaseModel):
    source_row_count: int
    filtered_row_count: int
    pivot_rows: int
    aggregate_rows: int
    details_rows: int
    pivot_preview: list[dict]
    aggregate_preview: list[dict]


class AnalyzeResponse(BaseModel):
    summary: str
    recommendations: list[str]
    model: str
    success: bool
    error: Optional[str] = None


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {
        "status": "ok",
        "version": "10.0.0",
        "engine_loaded": _engine is not None and _engine.has_data,
    }


@app.post("/ingest", response_model=IngestResponse)
async def ingest(
    file: UploadFile = File(...),
    sheet_name: Optional[str] = None,
):
    """Upload a file and get back column mapping + data preview."""
    global _engine

    content = await file.read()
    df = UniversalIngestor.from_buffer(content, filename=file.filename or "upload.xlsx", sheet_name=sheet_name)

    mapping = UniversalIngestor.auto_map_columns(list(df.columns))

    try:
        col_map = ColumnMap.auto_detect(list(df.columns))
    except ValueError:
        col_map = None

    _engine = PlannerEngine(col_map)
    _engine.load_master(df)

    return IngestResponse(
        rows=len(df),
        columns=list(df.columns),
        auto_mapping=mapping,
        preview=df.head(10).fillna("").to_dict(orient="records"),
    )


@app.post("/ingest/set-columns")
async def set_columns(col_map: ColumnMapRequest):
    """Manually set column mapping after ingest."""
    global _engine
    if _engine is None or not _engine.has_data:
        raise HTTPException(status_code=400, detail="No data loaded. Call /ingest first.")

    _engine._col_map = ColumnMap(
        person=col_map.person,
        created=col_map.created,
        completed=col_map.completed,
        task=col_map.task,
    )
    return {"status": "ok", "mapping": col_map.model_dump()}


@app.post("/merge", response_model=MergeResponse)
async def merge(
    file: UploadFile = File(...),
    sheet_name: Optional[str] = None,
):
    """Merge a new-format file into the loaded master."""
    global _engine
    if _engine is None or not _engine.has_data:
        raise HTTPException(status_code=400, detail="No master loaded. Call /ingest first.")

    content = await file.read()
    new_df = UniversalIngestor.from_buffer(content, filename=file.filename or "upload.xlsx", sheet_name=sheet_name)

    stats = _engine.merge_new(new_df)
    return MergeResponse(**stats)


@app.post("/run", response_model=RunResponse)
async def run():
    """Execute the aggregation pipeline on loaded data."""
    global _engine, _last_result
    if _engine is None or not _engine.has_data:
        raise HTTPException(status_code=400, detail="No data loaded.")

    result = _engine.run()
    _last_result = result

    return RunResponse(
        source_row_count=result.source_row_count,
        filtered_row_count=result.filtered_row_count,
        pivot_rows=len(result.pivot),
        aggregate_rows=len(result.aggregate),
        details_rows=len(result.details),
        pivot_preview=result.pivot.head(20).fillna("").to_dict(orient="records"),
        aggregate_preview=result.aggregate.head(20).fillna("").to_dict(orient="records"),
    )


@app.post("/export/excel")
async def export_excel():
    """Generate and download the formatted Excel report."""
    global _last_result
    if _last_result is None:
        raise HTTPException(status_code=400, detail="No results. Call /run first.")

    xlsx_bytes = ExcelReportWriter.write_to_buffer(
        pivot_df=_last_result.pivot,
        aggr_df=_last_result.aggregate,
        details_df=_last_result.details,
    )

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=planner_report.xlsx"},
    )


@app.post("/analyze", response_model=AnalyzeResponse)
async def analyze():
    """Run AI analysis on the last engine result."""
    global _last_result
    if _last_result is None:
        raise HTTPException(status_code=400, detail="No results. Call /run first.")

    from ..ai.analyst import Analyst

    analyst = Analyst()
    result = analyst.analyze(_last_result.aggregate, _last_result.pivot)

    return AnalyzeResponse(
        summary=result.summary,
        recommendations=result.recommendations,
        model=result.model,
        success=result.success,
        error=result.error,
    )
