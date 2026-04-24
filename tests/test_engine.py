"""Unit tests for PlannerEngine core logic."""

import pandas as pd
import pytest

from src.engine.processor import ColumnMap, PlannerEngine
from src.engine.ingestor import UniversalIngestor, smart_datetime, norm_text
from src.engine.excel_writer import ExcelReportWriter


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture
def sample_df():
    return pd.DataFrame({
        "Zasobnik": ["Kowalski", "Kowalski", "Nowak", "Nowak", "Nowak"],
        "Data utworzenia": ["01.01.2025", "15.02.2025", "01.03.2025", "10.04.2025", "20.01.2025"],
        "Data ukończenia": ["15.03.2025", "", "10.06.2025", "", ""],
    })


@pytest.fixture
def col_map():
    return ColumnMap(person="Zasobnik", created="Data utworzenia", completed="Data ukończenia")


# ── Ingestor tests ───────────────────────────────────────────────────────────

class TestIngestor:
    def test_smart_datetime_pl_format(self):
        s = pd.Series(["01.01.2025", "15.03.2025", ""])
        result = smart_datetime(s)
        assert pd.notna(result.iloc[0])
        assert pd.notna(result.iloc[1])
        assert pd.isna(result.iloc[2])

    def test_norm_text(self):
        s = pd.Series(["  Hello  World  ", "nan", "\u200bTest"])
        result = norm_text(s)
        assert result.iloc[0] == "Hello World"
        assert result.iloc[1] == ""
        assert result.iloc[2] == "Test"

    def test_auto_map_columns(self):
        cols = ["Zasobnik", "Data utworzenia", "Data ukończenia", "Nazwa zadania"]
        mapping = UniversalIngestor.auto_map_columns(cols)
        assert mapping["person"] == "Zasobnik"
        assert mapping["created"] == "Data utworzenia"
        assert mapping["completed"] == "Data ukończenia"

    def test_from_raw_text_tab(self):
        text = "Name\tDate\nAlice\t2025-01-01\nBob\t2025-02-01"
        df = UniversalIngestor.from_raw_text(text)
        assert len(df) == 2
        assert list(df.columns) == ["Name", "Date"]

    def test_from_raw_text_csv(self):
        text = "Name;Date\nAlice;2025-01-01"
        df = UniversalIngestor.from_raw_text(text)
        assert len(df) == 1


# ── Engine tests ─────────────────────────────────────────────────────────────

class TestEngine:
    def test_load_master(self, sample_df, col_map):
        engine = PlannerEngine(col_map)
        engine.load_master(sample_df)
        assert engine.has_data
        assert len(engine.df) == 5

    def test_run_produces_result(self, sample_df, col_map):
        engine = PlannerEngine(col_map)
        engine.load_master(sample_df)
        result = engine.run()
        assert result.source_row_count == 5
        assert len(result.pivot) > 0
        assert "Ukończone" in result.pivot.columns
        assert "Nieukończone" in result.pivot.columns

    def test_pivot_counts_correct(self, sample_df, col_map):
        engine = PlannerEngine(col_map)
        engine.load_master(sample_df)
        result = engine.run()
        total_done = result.pivot["Ukończone"].sum()
        total_undone = result.pivot["Nieukończone"].sum()
        # 2 tasks have completion dates, 3 do not (but one has empty created so may drop)
        assert total_done == 2
        assert total_undone >= 2

    def test_aggregate_columns(self, sample_df, col_map):
        engine = PlannerEngine(col_map)
        engine.load_master(sample_df)
        result = engine.run()
        assert "zadania_ukończone" in result.aggregate.columns
        assert "zadania_nieukończone" in result.aggregate.columns

    def test_details_statuses(self, sample_df, col_map):
        engine = PlannerEngine(col_map)
        engine.load_master(sample_df)
        result = engine.run()
        statuses = set(result.details["status"].unique())
        assert statuses <= {"ukończone", "nieukończone"}

    def test_auto_detect_columns(self):
        df = pd.DataFrame({
            "Zasobnik": ["A"], "Data utworzenia": ["01.01.2025"], "Data ukończenia": [""]
        })
        cm = ColumnMap.auto_detect(list(df.columns))
        assert cm.person == "Zasobnik"


# ── Excel writer tests ───────────────────────────────────────────────────────

class TestExcelWriter:
    def test_write_to_buffer(self, sample_df, col_map):
        engine = PlannerEngine(col_map)
        engine.load_master(sample_df)
        result = engine.run()
        xlsx_bytes = ExcelReportWriter.write_to_buffer(
            pivot_df=result.pivot,
            aggr_df=result.aggregate,
            details_df=result.details,
        )
        assert len(xlsx_bytes) > 0
        assert xlsx_bytes[:2] == b"PK"  # ZIP/XLSX magic bytes
