"""
PlannerEngine v10.0 — Pure headless data processing engine.

All aggregation, merging, filtering, and hierarchical pivot logic lives here.
Zero UI dependencies. Operates on DataFrames or raw buffers via UniversalIngestor.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd

from .ingestor import UniversalIngestor, best_guess_column, norm_text, smart_datetime


# ── Configuration ────────────────────────────────────────────────────────────

SENTINEL_CREATED = datetime(2025, 4, 10).date()  # "fake" date from Panasonic new system
MERGE_DATE_CUTOFF = datetime(2025, 4, 12).date()  # only new-format rows after this date

EXCLUDED_BINS = frozenset({"TO DO", "INSTRUKCJA", "HOLD"})


# ── Data containers ──────────────────────────────────────────────────────────

@dataclass
class ColumnMap:
    """Maps logical roles to actual DataFrame column names."""
    person: str
    created: str
    completed: str
    task: Optional[str] = None

    @classmethod
    def auto_detect(cls, columns: list[str]) -> "ColumnMap":
        mapping = UniversalIngestor.auto_map_columns(columns)
        if not mapping["person"] or not mapping["created"] or not mapping["completed"]:
            raise ValueError(
                f"Cannot auto-detect required columns from: {columns}. "
                "Provide an explicit ColumnMap."
            )
        return cls(
            person=mapping["person"],
            created=mapping["created"],
            completed=mapping["completed"],
            task=mapping["task"],
        )


@dataclass
class EngineResult:
    """Holds all outputs from a single engine run."""
    pivot: pd.DataFrame          # Osoba × Rok × Kwartał → Ukończone / Nieukończone
    aggregate: pd.DataFrame      # Tabular summary per zasobnik/rok/kwartał
    details: pd.DataFrame        # Flat task list with status
    source_row_count: int = 0
    filtered_row_count: int = 0
    merge_stats: dict = field(default_factory=dict)


# ── Engine ───────────────────────────────────────────────────────────────────

class PlannerEngine:
    """
    Stateless data processing engine.

    Usage:
        engine = PlannerEngine(col_map)
        engine.load_master(df_or_path)
        engine.merge_new(df_or_path)       # optional
        result = engine.run()
    """

    def __init__(self, col_map: Optional[ColumnMap] = None):
        self._col_map: Optional[ColumnMap] = col_map
        self._master: Optional[pd.DataFrame] = None
        self._merged: Optional[pd.DataFrame] = None

    # ── Properties ───────────────────────────────────────────────────────

    @property
    def col_map(self) -> ColumnMap:
        if self._col_map is None:
            raise RuntimeError("ColumnMap not set. Load data first or provide explicitly.")
        return self._col_map

    @property
    def df(self) -> pd.DataFrame:
        """The active working DataFrame (merged if available, otherwise master)."""
        if self._merged is not None:
            return self._merged
        if self._master is not None:
            return self._master
        raise RuntimeError("No data loaded.")

    @property
    def has_data(self) -> bool:
        return self._master is not None

    # ── Data loading ─────────────────────────────────────────────────────

    def load_master(self, source: pd.DataFrame) -> "PlannerEngine":
        """Load the master (archival) dataset."""
        self._master = source.copy()
        self._merged = None
        if self._col_map is None:
            self._col_map = ColumnMap.auto_detect(list(source.columns))
        return self

    def merge_new(self, new_df: pd.DataFrame) -> dict:
        """
        Merge a new-format DataFrame into master using unique-key dedup.
        Returns merge statistics.
        """
        if self._master is None:
            raise RuntimeError("Load master data first.")

        cm = self.col_map

        # Date filter: only keep rows created after the cutoff
        created_dt = smart_datetime(new_df[cm.created]).dt.date
        mask_keep = created_dt > MERGE_DATE_CUTOFF
        filtered = new_df.loc[mask_keep].copy()

        # Build unique keys
        key_master = self._build_unique_key(self._master)
        key_new = self._build_unique_key(filtered)

        master_keys = set(key_master.dropna().tolist())
        filtered["__key__"] = key_new
        new_only = filtered[~filtered["__key__"].isin(master_keys)].drop(
            columns="__key__", errors="ignore"
        )

        self._merged = pd.concat([self._master, new_only], ignore_index=True)

        stats = {
            "new_file_rows": len(new_df),
            "after_date_filter": len(filtered),
            "duplicates_skipped": len(filtered) - len(new_only),
            "rows_added": len(new_only),
            "total_rows": len(self._merged),
        }
        return stats

    # ── Core processing ──────────────────────────────────────────────────

    def run(self) -> EngineResult:
        """Execute full aggregation pipeline. Returns EngineResult."""
        df = self.df
        cm = self.col_map

        person, created, completed, keep_mask = self._prepare_series(df, cm)

        pivot = self._build_pivot(person, created, completed)
        aggregate = self._build_aggregate(person, created, completed)
        details = self._build_details(person, created, completed)

        return EngineResult(
            pivot=pivot,
            aggregate=aggregate,
            details=details,
            source_row_count=len(df),
            filtered_row_count=int(keep_mask.sum()),
        )

    # ── Pivot: Osoba > Rok > Kwartał → Ukończone / Nieukończone ─────────

    @staticmethod
    def _build_pivot(
        person: pd.Series,
        created: pd.Series,
        completed: pd.Series,
    ) -> pd.DataFrame:
        has_completed = completed.notna()
        not_completed = completed.isna() & created.notna()

        # Ukończone — grouped by completion date
        done_df = (
            pd.DataFrame(
                {
                    "osoba": person[has_completed],
                    "rok": completed[has_completed].dt.year,
                    "kwartał": completed[has_completed].dt.quarter,
                }
            )
            .dropna(subset=["osoba", "rok", "kwartał"])
            .groupby(["osoba", "rok", "kwartał"])
            .size()
            .rename("Ukończone")
            .reset_index()
        )

        # Nieukończone — grouped by creation date
        undone_df = (
            pd.DataFrame(
                {
                    "osoba": person[not_completed],
                    "rok": created[not_completed].dt.year,
                    "kwartał": created[not_completed].dt.quarter,
                }
            )
            .dropna(subset=["osoba", "rok", "kwartał"])
            .groupby(["osoba", "rok", "kwartał"])
            .size()
            .rename("Nieukończone")
            .reset_index()
        )

        pivot = pd.merge(done_df, undone_df, on=["osoba", "rok", "kwartał"], how="outer")
        pivot["Ukończone"] = pivot["Ukończone"].fillna(0).astype(int)
        pivot["Nieukończone"] = pivot["Nieukończone"].fillna(0).astype(int)
        pivot = pivot.sort_values(
            by=["osoba", "rok", "kwartał"],
            key=lambda s: s.str.casefold() if s.name == "osoba" else s,
        ).reset_index(drop=True)
        return pivot

    # ── Aggregate: Rok × Kwartał × Zasobnik ─────────────────────────────

    @staticmethod
    def _build_aggregate(
        person: pd.Series,
        created: pd.Series,
        completed: pd.Series,
    ) -> pd.DataFrame:
        # Ukończone by completion date
        done = (
            pd.DataFrame(
                {
                    "rok": completed.dt.year,
                    "kwartał": completed.dt.quarter,
                    "zasobnik": person,
                }
            )
            .dropna(subset=["rok", "kwartał"])
            .groupby(["rok", "kwartał", "zasobnik"], dropna=False)
            .size()
            .rename("zadania_ukończone")
            .reset_index()
        )

        # Nieukończone by creation date (where completed is NaN)
        not_done_mask = completed.isna()
        not_done_created = created.where(not_done_mask)
        not_done = (
            pd.DataFrame(
                {
                    "rok": not_done_created.dt.year,
                    "kwartał": not_done_created.dt.quarter,
                    "zasobnik": person.where(not_done_mask),
                }
            )
            .dropna(subset=["rok", "kwartał", "zasobnik"])
            .groupby(["rok", "kwartał", "zasobnik"], dropna=False)
            .size()
            .rename("zadania_nieukończone")
            .reset_index()
        )

        out = pd.merge(done, not_done, on=["rok", "kwartał", "zasobnik"], how="outer")
        out["zadania_ukończone"] = out["zadania_ukończone"].fillna(0).astype(int)
        out["zadania_nieukończone"] = out["zadania_nieukończone"].fillna(0).astype(int)
        out.sort_values(
            by=["zasobnik", "rok", "kwartał"],
            key=lambda col: col.str.casefold() if col.name == "zasobnik" else col,
            inplace=True,
        )
        out.reset_index(drop=True, inplace=True)
        return out

    # ── Details: flat task list ──────────────────────────────────────────

    @staticmethod
    def _build_details(
        person: pd.Series,
        created: pd.Series,
        completed: pd.Series,
    ) -> pd.DataFrame:
        has_done = completed.notna()
        not_done_mask = completed.isna()

        det_done = pd.DataFrame(
            {
                "osoba": person[has_done],
                "data_utworzenia": created[has_done],
                "data_ukonczenia": completed[has_done],
                "status": "ukończone",
                "rok": completed[has_done].dt.year,
                "kwartał": completed[has_done].dt.quarter,
            }
        ).dropna(subset=["rok", "kwartał"])

        det_not_done = pd.DataFrame(
            {
                "osoba": person[not_done_mask],
                "data_utworzenia": created[not_done_mask],
                "data_ukonczenia": pd.NaT,
                "status": "nieukończone",
                "rok": created[not_done_mask].dt.year,
                "kwartał": created[not_done_mask].dt.quarter,
            }
        ).dropna(subset=["osoba", "rok", "kwartał"])

        details = (
            pd.concat([det_done, det_not_done], ignore_index=True)
            .sort_values(
                by=["osoba", "rok", "kwartał", "data_utworzenia", "data_ukonczenia"],
                key=lambda col: col.str.casefold() if col.name == "osoba" else col,
            )
            .reset_index(drop=True)
        )
        return details

    # ── Internal helpers ─────────────────────────────────────────────────

    @staticmethod
    def _filter_bins(person: pd.Series) -> pd.Series:
        return ~person.str.upper().isin(EXCLUDED_BINS)

    def _prepare_series(
        self, df: pd.DataFrame, cm: ColumnMap
    ) -> tuple[pd.Series, pd.Series, pd.Series, pd.Series]:
        """Normalize and filter the three core series. Returns (person, created, completed, keep_mask)."""
        person = norm_text(df[cm.person]).replace("", "—")
        keep = self._filter_bins(person)
        person = person[keep]
        created = smart_datetime(df[cm.created])[keep]
        completed = smart_datetime(df[cm.completed])[keep]
        return person, created, completed, keep

    def _build_unique_key(self, df: pd.DataFrame) -> pd.Series:
        """
        Unique key for deduplication.
        Priority: ID column > (person, task, created, completed) composite.
        Sentinel date logic: if created == SENTINEL, omit it from key and use completed instead.
        """
        cm = self.col_map
        cols = list(df.columns)
        id_col = "Identyfikator zadania"

        if id_col in cols:
            return norm_text(df[id_col])

        person = norm_text(df.get(cm.person, pd.Series([""] * len(df)))).str.lower()
        task = (
            norm_text(df.get(cm.task, pd.Series([""] * len(df)))).str.lower()
            if cm.task
            else pd.Series([""] * len(df))
        )
        created_dt = smart_datetime(df.get(cm.created, pd.Series([""] * len(df))))
        completed_dt = smart_datetime(df.get(cm.completed, pd.Series([""] * len(df))))

        created_date = created_dt.dt.date
        completed_date = completed_dt.dt.date

        created_str = created_date.astype(str).where(pd.notna(created_date), "")
        completed_str = completed_date.astype(str).where(pd.notna(completed_date), "")

        use_created = ~created_date.eq(SENTINEL_CREATED)

        key = (
            person
            + "||"
            + task.fillna("")
            + "||"
            + created_str.where(use_created, "")
            + "||"
            + completed_str.where(use_created, completed_str)
        )
        return key
