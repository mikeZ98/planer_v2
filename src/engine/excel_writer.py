"""
ExcelReportWriter v10.0 — Business-critical xlsxwriter formatting logic.

Generates multi-sheet Excel reports with:
  - "Agregat" sheet: tabular summary (Rok, Kwartał, Zasobnik, Ukończone, Nieukończone)
  - "Szczegóły" sheet: flat task list with status
  - "Pivot" sheet: hierarchical layout for multi-level category axis
  - "Wykres" sheet: charts with [Osoba] > [Rok] > [KW] X-axis hierarchy
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Optional, Union

import pandas as pd


class ExcelReportWriter:
    """
    Writes the final formatted Excel workbook with charts.
    Preserves the xlsxwriter multi-level category axis trick for hierarchical grouping.
    """

    # ── Public API ───────────────────────────────────────────────────────

    @classmethod
    def write(
        cls,
        dest: Union[str, Path, io.BytesIO],
        pivot_df: pd.DataFrame,
        aggr_df: Optional[pd.DataFrame] = None,
        details_df: Optional[pd.DataFrame] = None,
        title: str = "Ukończone vs Nieukończone (Osoba / Rok / Kwartał)",
    ) -> None:
        """Write the full report to *dest* (file path or BytesIO buffer)."""
        with pd.ExcelWriter(dest, engine="xlsxwriter") as writer:
            book = writer.book

            if aggr_df is not None:
                aggr_df.to_excel(writer, sheet_name="Agregat", index=False)
                cls._format_table_sheet(writer, book, "Agregat", aggr_df)

            if details_df is not None:
                details_df.to_excel(writer, sheet_name="Szczegóły", index=False)
                cls._format_table_sheet(writer, book, "Szczegóły", details_df)

            cls._write_pivot_sheet(writer, book, pivot_df)
            cls._write_chart_sheet(writer, book, pivot_df, title)

    @classmethod
    def write_to_buffer(
        cls,
        pivot_df: pd.DataFrame,
        aggr_df: Optional[pd.DataFrame] = None,
        details_df: Optional[pd.DataFrame] = None,
        title: str = "Ukończone vs Nieukończone (Osoba / Rok / Kwartał)",
    ) -> bytes:
        """Write the report into memory and return raw bytes (for web download)."""
        buf = io.BytesIO()
        cls.write(buf, pivot_df, aggr_df, details_df, title)
        return buf.getvalue()

    # ── Pivot sheet ──────────────────────────────────────────────────────

    @classmethod
    def _write_pivot_sheet(
        cls,
        writer: pd.ExcelWriter,
        book,
        pivot_df: pd.DataFrame,
    ) -> None:
        ws = book.add_worksheet("Pivot")
        writer.sheets["Pivot"] = ws

        cat_person = pivot_df["osoba"].tolist()
        cat_year = pivot_df["rok"].astype(int).tolist()
        cat_quart = pivot_df["kwartał"].astype(int).map(lambda q: f"KW{q}").tolist()

        # Rows 0-2: hierarchical X-axis  [Osoba] > [Rok] > [Kwartał]
        ws.write_row(0, 1, cat_person)
        ws.write_row(1, 1, cat_year)
        ws.write_row(2, 1, cat_quart)

        # Column A headers
        header_fmt = book.add_format({"bold": True, "bg_color": "#D9E1F2"})
        ws.write(0, 0, "Osoba", header_fmt)
        ws.write(1, 0, "Rok", header_fmt)
        ws.write(2, 0, "Kwartał", header_fmt)

        # Data rows (chart series source)
        data_fmt = book.add_format({"bold": True})
        ws.write(3, 0, "Ukończone", data_fmt)
        ws.write_row(3, 1, pivot_df["Ukończone"].tolist())
        ws.write(4, 0, "Nieukończone", data_fmt)
        ws.write_row(4, 1, pivot_df["Nieukończone"].tolist())

        # Quarterly summary section (no person breakdown)
        if not pivot_df.empty:
            per_q = (
                pivot_df.groupby(["rok", "kwartał"], as_index=False)[
                    ["Ukończone", "Nieukończone"]
                ]
                .sum()
                .sort_values(["rok", "kwartał"])
            )
            q_labels = [
                f"{int(r)} | KW{int(k)}"
                for r, k in zip(per_q["rok"], per_q["kwartał"])
            ]
            ws.write_row(7, 0, ["Kwartał (oś)", *q_labels])
            ws.write(9, 0, "Ukończone (suma)")
            ws.write_row(9, 1, per_q["Ukończone"].tolist())
            ws.write(10, 0, "Nieukończone (suma)")
            ws.write_row(10, 1, per_q["Nieukończone"].tolist())

        # Auto-fit column A
        ws.set_column(0, 0, 20)

    # ── Chart sheet ──────────────────────────────────────────────────────

    @classmethod
    def _write_chart_sheet(
        cls,
        writer: pd.ExcelWriter,
        book,
        pivot_df: pd.DataFrame,
        title: str,
    ) -> None:
        ws = book.add_worksheet("Wykres")
        writer.sheets["Wykres"] = ws

        last_col = len(pivot_df)
        if last_col == 0:
            ws.write(0, 0, "Brak danych do wykresu.")
            return

        # ── Chart 1: Multi-level Category Axis (Person > Year > Quarter) ──
        chart1 = book.add_chart({"type": "column"})

        # categories span rows 0-2 → xlsxwriter creates hierarchical X-axis
        chart1.add_series(
            {
                "name": "='Pivot'!$A$4",  # Ukończone
                "categories": ["Pivot", 0, 1, 2, last_col],
                "values": ["Pivot", 3, 1, 3, last_col],
                "data_labels": {"value": True, "font": {"size": 8}},
                "fill": {"color": "#43A047"},  # green
                "border": {"color": "#2E7D32"},
            }
        )
        chart1.add_series(
            {
                "name": "='Pivot'!$A$5",  # Nieukończone
                "categories": ["Pivot", 0, 1, 2, last_col],
                "values": ["Pivot", 4, 1, 4, last_col],
                "data_labels": {"value": True, "font": {"size": 8}},
                "fill": {"color": "#E53935"},  # red
                "border": {"color": "#B71C1C"},
            }
        )
        chart1.set_title({"name": title})
        chart1.set_legend({"position": "top"})
        chart1.set_y_axis({"major_gridlines": {"visible": True}})
        chart1.set_size({"x_scale": 2.2, "y_scale": 1.6})
        ws.insert_chart("B2", chart1)

        # ── Chart 2: Quarterly summary (clustered, no person breakdown) ──
        if not pivot_df.empty:
            per_q = (
                pivot_df.groupby(["rok", "kwartał"], as_index=False)[
                    ["Ukończone", "Nieukończone"]
                ]
                .sum()
                .sort_values(["rok", "kwartał"])
            )
            last_col_q = len(per_q)
            if last_col_q > 0:
                chart2 = book.add_chart({"type": "column"})
                chart2.add_series(
                    {
                        "name": "='Pivot'!$A$10",
                        "categories": ["Pivot", 7, 1, 7, last_col_q],
                        "values": ["Pivot", 9, 1, 9, last_col_q],
                        "data_labels": {"value": True},
                        "fill": {"color": "#43A047"},
                        "border": {"color": "#2E7D32"},
                    }
                )
                chart2.add_series(
                    {
                        "name": "='Pivot'!$A$11",
                        "categories": ["Pivot", 7, 1, 7, last_col_q],
                        "values": ["Pivot", 10, 1, 10, last_col_q],
                        "data_labels": {"value": True},
                        "fill": {"color": "#E53935"},
                        "border": {"color": "#B71C1C"},
                    }
                )
                chart2.set_title(
                    {"name": "Suma kwartalna — Ukończone vs Nieukończone"}
                )
                chart2.set_legend({"position": "top"})
                chart2.set_y_axis({"major_gridlines": {"visible": True}})
                chart2.set_size({"x_scale": 1.6, "y_scale": 1.6})
                ws.insert_chart("N2", chart2)

    # ── Table formatting helper ──────────────────────────────────────────

    @staticmethod
    def _format_table_sheet(
        writer: pd.ExcelWriter,
        book,
        sheet_name: str,
        df: pd.DataFrame,
    ) -> None:
        ws = writer.sheets[sheet_name]
        header_fmt = book.add_format(
            {
                "bold": True,
                "bg_color": "#1E1E2E",
                "font_color": "#CDD6F4",
                "border": 1,
            }
        )
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, header_fmt)
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            ws.set_column(col_idx, col_idx, min(max_len + 4, 40))
