"""
Universal Ingestor v10.0
Supports: .xlsx (multi-sheet), .xls, .xlsm, .csv, .json, .txt (structured paste).
Returns a normalized pandas DataFrame regardless of source format.
"""

from __future__ import annotations

import io
import json
import unicodedata
from pathlib import Path
from typing import BinaryIO, Optional, Union

import pandas as pd


# ── Column auto-detection ────────────────────────────────────────────────────

def _norm_token(s: str) -> str:
    s = str(s)
    s = "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )
    s = s.lower()
    for ch in ["-", "_", " ", "/", "\\", ".", "(", ")", "[", "]", ":", ";", "|"]:
        s = s.replace(ch, "")
    return s


def best_guess_column(columns: list[str], candidates: list[str]) -> Optional[str]:
    """Fuzzy-match a column name from *columns* against a list of *candidates*."""
    cols_norm = {col: _norm_token(col) for col in columns}
    cand_norm = [_norm_token(c) for c in candidates]

    for col, ncol in cols_norm.items():
        if ncol in cand_norm:
            return col

    best, best_len = None, 0
    for col, ncol in cols_norm.items():
        for c in cand_norm:
            if c in ncol and len(c) > best_len:
                best, best_len = col, len(c)
    return best or (columns[0] if columns else None)


def detect_sheet(path: Union[str, Path]) -> Optional[str]:
    """Pick the most likely data sheet from an Excel workbook."""
    try:
        xls = pd.ExcelFile(path)
        preferred = [
            s
            for s in xls.sheet_names
            if str(s).strip().lower()
            in {"data", "dane", "tasks", "zadania", "sheet1", "arkusz1"}
        ]
        return preferred[0] if preferred else xls.sheet_names[0]
    except Exception:
        return None


# ── Norm helpers ─────────────────────────────────────────────────────────────

def norm_text(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.replace("\u200b", "", regex=False)
        .replace({"nan": "", "None": ""})
    )


def smart_datetime(s: pd.Series) -> pd.Series:
    """Robust date parser — handles PL dd.mm.yyyy, ISO, and mixed formats."""
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    mask = dt.isna() & s.astype(str).str.contains(r"\d{1,2}\.\d{1,2}\.\d{2,4}")
    if mask.any():
        vals = s[mask].astype(str).str.strip()
        dt2 = pd.to_datetime(vals, format="%d.%m.%Y", errors="coerce")
        still = dt2.isna()
        if still.any():
            dt2.loc[still] = pd.to_datetime(
                vals[still], format="%d.%m.%Y %H:%M", errors="coerce"
            )
        dt.loc[mask] = dt2
    return dt


# ── Universal Ingestor ───────────────────────────────────────────────────────

class UniversalIngestor:
    """
    Ingests data from any supported format and returns a clean DataFrame.

    Supported formats:
      - .xlsx / .xls / .xlsm  (auto-detects best sheet)
      - .csv                   (auto-detects delimiter)
      - .json                  (records or columnar orientation)
      - .txt                   (tab / semicolon / comma delimited)
      - raw bytes / StringIO   (for web upload buffers)
    """

    EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
    CSV_EXTENSIONS = {".csv"}
    JSON_EXTENSIONS = {".json"}
    TEXT_EXTENSIONS = {".txt", ".tsv"}

    # ── Public API ───────────────────────────────────────────────────────

    @classmethod
    def from_file(
        cls,
        path: Union[str, Path],
        sheet_name: Optional[str] = None,
    ) -> pd.DataFrame:
        """Ingest from a filesystem path.  Auto-detects format by extension."""
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        ext = path.suffix.lower()
        if ext in cls.EXCEL_EXTENSIONS:
            return cls._read_excel(path, sheet_name)
        elif ext in cls.CSV_EXTENSIONS:
            return cls._read_csv(path)
        elif ext in cls.JSON_EXTENSIONS:
            return cls._read_json(path)
        elif ext in cls.TEXT_EXTENSIONS:
            return cls._read_text(path)
        else:
            raise ValueError(f"Unsupported file extension: {ext}")

    @classmethod
    def from_buffer(
        cls,
        buf: Union[BinaryIO, bytes, str],
        filename: str = "upload.xlsx",
        sheet_name: Optional[str] = None,
    ) -> pd.DataFrame:
        """Ingest from an in-memory buffer (web upload, API payload, etc.)."""
        ext = Path(filename).suffix.lower()

        if isinstance(buf, str):
            return cls._parse_raw_text(buf)
        if isinstance(buf, bytes):
            buf = io.BytesIO(buf)

        if ext in cls.EXCEL_EXTENSIONS:
            return cls._read_excel(buf, sheet_name)
        elif ext in cls.CSV_EXTENSIONS:
            raw = buf.read().decode("utf-8", errors="replace")
            return cls._parse_csv_text(raw)
        elif ext in cls.JSON_EXTENSIONS:
            raw = buf.read().decode("utf-8", errors="replace")
            return cls._parse_json_text(raw)
        else:
            raw = buf.read().decode("utf-8", errors="replace")
            return cls._parse_raw_text(raw)

    @classmethod
    def from_raw_text(cls, text: str) -> pd.DataFrame:
        """Ingest from raw pasted text (tab/csv structured)."""
        return cls._parse_raw_text(text)

    # ── Column mapping helpers ───────────────────────────────────────────

    @staticmethod
    def auto_map_columns(columns: list[str]) -> dict[str, Optional[str]]:
        """
        Auto-detect standard column roles from a DataFrame's column list.
        Returns dict with keys: person, created, completed, task.
        """
        return {
            "person": best_guess_column(
                columns,
                [
                    "zasobnik", "osoba", "assignee", "owner",
                    "wykonawca", "przypisane do", "assigned to", "user",
                ],
            ),
            "created": best_guess_column(
                columns,
                [
                    "data utworzenia", "utworzenia", "created",
                    "creation date", "created at", "start date",
                ],
            ),
            "completed": best_guess_column(
                columns,
                [
                    "data ukończenia", "ukonczenia", "completed",
                    "done", "closed", "end date", "resolution date",
                ],
            ),
            "task": best_guess_column(
                columns,
                ["nazwa zadania", "tytuł", "title", "task", "nazwa"],
            ),
        }

    # ── Private readers ──────────────────────────────────────────────────

    @classmethod
    def _read_excel(
        cls,
        source: Union[Path, BinaryIO],
        sheet_name: Optional[str] = None,
    ) -> pd.DataFrame:
        if sheet_name is None and isinstance(source, Path):
            sheet_name = detect_sheet(source)
        df = pd.read_excel(source, sheet_name=sheet_name or 0)
        return df

    @classmethod
    def _read_csv(cls, path: Path) -> pd.DataFrame:
        raw = path.read_text(encoding="utf-8", errors="replace")
        return cls._parse_csv_text(raw)

    @classmethod
    def _read_json(cls, path: Path) -> pd.DataFrame:
        raw = path.read_text(encoding="utf-8", errors="replace")
        return cls._parse_json_text(raw)

    @classmethod
    def _read_text(cls, path: Path) -> pd.DataFrame:
        raw = path.read_text(encoding="utf-8", errors="replace")
        return cls._parse_raw_text(raw)

    # ── Private parsers ──────────────────────────────────────────────────

    @staticmethod
    def _parse_csv_text(raw: str) -> pd.DataFrame:
        for sep in [";", "\t", ",", "|"]:
            try:
                df = pd.read_csv(io.StringIO(raw), sep=sep)
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
        return pd.read_csv(io.StringIO(raw))

    @staticmethod
    def _parse_json_text(raw: str) -> pd.DataFrame:
        data = json.loads(raw)
        if isinstance(data, list):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            if all(isinstance(v, list) for v in data.values()):
                return pd.DataFrame(data)
            return pd.json_normalize(data)
        raise ValueError("Unsupported JSON structure")

    @staticmethod
    def _parse_raw_text(raw: str) -> pd.DataFrame:
        lines = [l for l in raw.strip().splitlines() if l.strip()]
        if not lines:
            return pd.DataFrame()

        for sep in ["\t", ";", ",", "|"]:
            cols_first = lines[0].split(sep)
            if len(cols_first) > 1:
                header = [c.strip() for c in cols_first]
                rows = []
                for line in lines[1:]:
                    parts = [c.strip() for c in line.split(sep)]
                    if len(parts) == len(header):
                        rows.append(parts)
                    else:
                        padded = parts + [""] * (len(header) - len(parts))
                        rows.append(padded[: len(header)])
                return pd.DataFrame(rows, columns=header)

        return pd.DataFrame({"raw": lines})
