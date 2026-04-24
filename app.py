"""
Planner v10.0 — Streamlit Dark-Mode Dashboard

Three views:
  1. Ingest  — Drag & drop multiple files + raw text area
  2. Analytics — Interactive Plotly charts (hierarchical X-axis)
  3. Export  — Generate formatted Excel + PDF reports

Launch:  uv run streamlit run app.py --server.port 8501
"""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.engine.processor import ColumnMap, EngineResult, PlannerEngine
from src.engine.ingestor import UniversalIngestor
from src.engine.excel_writer import ExcelReportWriter

# ── Page config ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Planner v10 — Data Engineering Service",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Dark theme CSS ───────────────────────────────────────────────────────────

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Inter:wght@400;500;600;700&display=swap');

    :root {
        --bg-primary: #0f0f1a;
        --bg-secondary: #1a1a2e;
        --bg-card: #16213e;
        --accent-green: #43A047;
        --accent-red: #E53935;
        --accent-blue: #1E88E5;
        --text-primary: #e0e0e0;
        --text-muted: #8892b0;
        --border: #233554;
    }

    .stApp {
        background-color: var(--bg-primary);
        color: var(--text-primary);
        font-family: 'Inter', sans-serif;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: var(--bg-secondary);
        border-radius: 12px;
        padding: 4px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: 600;
    }

    .metric-card {
        background: linear-gradient(135deg, var(--bg-card) 0%, var(--bg-secondary) 100%);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 20px;
        text-align: center;
    }

    .metric-card h3 {
        color: var(--text-muted);
        font-size: 0.85rem;
        margin-bottom: 4px;
        font-weight: 500;
    }

    .metric-card .value {
        font-size: 2rem;
        font-weight: 700;
        font-family: 'JetBrains Mono', monospace;
    }

    .metric-card .value.green { color: var(--accent-green); }
    .metric-card .value.red { color: var(--accent-red); }
    .metric-card .value.blue { color: var(--accent-blue); }

    div[data-testid="stSidebar"] {
        background-color: var(--bg-secondary);
    }

    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .status-badge.ok { background: #1b5e20; color: #a5d6a7; }
    .status-badge.warn { background: #e65100; color: #ffcc80; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── Session state initialization ─────────────────────────────────────────────

def _init_state():
    defaults = {
        "engine": None,
        "result": None,
        "col_map": None,
        "files_loaded": [],
        "merge_stats": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📊 Planner v10.0")
    st.caption("Universal Data Engineering Service")
    st.divider()

    engine: PlannerEngine | None = st.session_state.engine

    if engine and engine.has_data:
        st.markdown(
            '<span class="status-badge ok">Engine loaded</span>',
            unsafe_allow_html=True,
        )
        st.metric("Rows in dataset", len(engine.df))
    else:
        st.markdown(
            '<span class="status-badge warn">No data</span>',
            unsafe_allow_html=True,
        )

    st.divider()
    st.caption("Built for 2026. No limits.")


# ── Tabs ─────────────────────────────────────────────────────────────────────

tab_ingest, tab_analytics, tab_export = st.tabs(["📥 Ingest", "📈 Analytics", "📤 Export"])


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — INGEST
# ══════════════════════════════════════════════════════════════════════════════

with tab_ingest:
    st.header("Data Ingestion")

    col_upload, col_text = st.columns([1, 1])

    with col_upload:
        st.subheader("File Upload")
        uploaded_files = st.file_uploader(
            "Drag & drop files here",
            type=["xlsx", "xls", "xlsm", "csv", "json", "txt"],
            accept_multiple_files=True,
            help="Supports: Excel (multi-format), CSV, JSON, TXT (structured)",
        )

    with col_text:
        st.subheader("Raw Text Paste")
        raw_text = st.text_area(
            "Paste structured data (tab/csv/semicolon separated)",
            height=200,
            placeholder="Zasobnik\tData utworzenia\tData ukończenia\nKowalski\t01.01.2025\t15.03.2025",
        )

    # ── Process uploads ──────────────────────────────────────────────────

    if uploaded_files or raw_text.strip():
        frames: list[pd.DataFrame] = []

        if uploaded_files:
            for f in uploaded_files:
                try:
                    df = UniversalIngestor.from_buffer(f.read(), filename=f.name)
                    frames.append(df)
                    st.success(f"✓ **{f.name}** — {len(df)} rows, {len(df.columns)} columns")
                except Exception as e:
                    st.error(f"✗ **{f.name}** — {e}")

        if raw_text.strip():
            try:
                df = UniversalIngestor.from_raw_text(raw_text)
                frames.append(df)
                st.success(f"✓ **Raw text** — {len(df)} rows, {len(df.columns)} columns")
            except Exception as e:
                st.error(f"✗ Raw text — {e}")

        if frames:
            combined = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
            cols = list(combined.columns)
            mapping = UniversalIngestor.auto_map_columns(cols)

            st.divider()
            st.subheader("Column Mapping")

            c1, c2 = st.columns(2)
            with c1:
                person_col = st.selectbox(
                    "Zasobnik (osoba)", cols,
                    index=cols.index(mapping["person"]) if mapping["person"] in cols else 0,
                )
                created_col = st.selectbox(
                    "Data utworzenia", cols,
                    index=cols.index(mapping["created"]) if mapping["created"] in cols else 0,
                )
            with c2:
                completed_col = st.selectbox(
                    "Data ukończenia", cols,
                    index=cols.index(mapping["completed"]) if mapping["completed"] in cols else 0,
                )
                task_col = st.selectbox(
                    "Nazwa zadania (optional)", ["(none)"] + cols,
                    index=(cols.index(mapping["task"]) + 1) if mapping["task"] in cols else 0,
                )

            col_map = ColumnMap(
                person=person_col,
                created=created_col,
                completed=completed_col,
                task=task_col if task_col != "(none)" else None,
            )

            if st.button("🚀 Load as Master", type="primary", use_container_width=True):
                engine = PlannerEngine(col_map)
                engine.load_master(combined)
                st.session_state.engine = engine
                st.session_state.col_map = col_map
                st.session_state.files_loaded = [f.name for f in (uploaded_files or [])]
                st.success(f"Master loaded: **{len(combined)}** rows")
                st.rerun()

            # ── Merge additional file ────────────────────────────────────
            if st.session_state.engine and st.session_state.engine.has_data:
                st.divider()
                st.subheader("Merge Additional File")
                merge_file = st.file_uploader(
                    "Upload new-format file to merge",
                    type=["xlsx", "xls", "xlsm", "csv", "json", "txt"],
                    key="merge_upload",
                )
                if merge_file and st.button("🔀 Merge into Master"):
                    try:
                        merge_df = UniversalIngestor.from_buffer(
                            merge_file.read(), filename=merge_file.name
                        )
                        stats = st.session_state.engine.merge_new(merge_df)
                        st.session_state.merge_stats = stats
                        st.success(
                            f"Merged: **+{stats['rows_added']}** new rows "
                            f"({stats['duplicates_skipped']} duplicates skipped). "
                            f"Total: **{stats['total_rows']}** rows."
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Merge failed: {e}")

            # ── Data preview ─────────────────────────────────────────────
            st.divider()
            st.subheader("Data Preview")
            st.dataframe(combined.head(50), use_container_width=True, height=300)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════

with tab_analytics:
    st.header("Analytics Dashboard")

    engine = st.session_state.engine
    if engine is None or not engine.has_data:
        st.info("Load data in the **Ingest** tab first.")
    else:
        # Run engine
        if st.button("🔄 Run Analysis", type="primary", use_container_width=True):
            result = engine.run()
            st.session_state.result = result
            st.rerun()

        result: EngineResult | None = st.session_state.result

        if result is None:
            st.info("Click **Run Analysis** to process loaded data.")
        else:
            pivot = result.pivot
            aggregate = result.aggregate

            # ── Metric cards ─────────────────────────────────────────────
            total_done = int(aggregate["zadania_ukończone"].sum()) if len(aggregate) > 0 else 0
            total_undone = int(aggregate["zadania_nieukończone"].sum()) if len(aggregate) > 0 else 0
            total = total_done + total_undone
            rate = (total_done / max(total, 1)) * 100

            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(
                    f'<div class="metric-card"><h3>Total Tasks</h3>'
                    f'<div class="value blue">{total}</div></div>',
                    unsafe_allow_html=True,
                )
            with m2:
                st.markdown(
                    f'<div class="metric-card"><h3>Completed</h3>'
                    f'<div class="value green">{total_done}</div></div>',
                    unsafe_allow_html=True,
                )
            with m3:
                st.markdown(
                    f'<div class="metric-card"><h3>Incomplete</h3>'
                    f'<div class="value red">{total_undone}</div></div>',
                    unsafe_allow_html=True,
                )
            with m4:
                st.markdown(
                    f'<div class="metric-card"><h3>Completion Rate</h3>'
                    f'<div class="value {"green" if rate >= 70 else "red"}">{rate:.1f}%</div></div>',
                    unsafe_allow_html=True,
                )

            st.divider()

            # ── Chart 1: Hierarchical bar chart (Person > Year > Quarter) ─
            st.subheader("Ukończone vs Nieukończone — Osoba > Rok > Kwartał")

            if not pivot.empty:
                labels = [
                    f"{row['osoba']} | {int(row['rok'])} | KW{int(row['kwartał'])}"
                    for _, row in pivot.iterrows()
                ]

                fig1 = go.Figure()
                fig1.add_trace(
                    go.Bar(
                        x=labels,
                        y=pivot["Ukończone"],
                        name="Ukończone",
                        marker_color="#43A047",
                        text=pivot["Ukończone"],
                        textposition="auto",
                    )
                )
                fig1.add_trace(
                    go.Bar(
                        x=labels,
                        y=pivot["Nieukończone"],
                        name="Nieukończone",
                        marker_color="#E53935",
                        text=pivot["Nieukończone"],
                        textposition="auto",
                    )
                )
                fig1.update_layout(
                    barmode="group",
                    template="plotly_dark",
                    paper_bgcolor="#0f0f1a",
                    plot_bgcolor="#1a1a2e",
                    font=dict(family="Inter", color="#e0e0e0"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    height=500,
                    margin=dict(b=120),
                    xaxis=dict(tickangle=-45),
                )
                st.plotly_chart(fig1, use_container_width=True)

            # ── Chart 2: Quarterly summary ───────────────────────────────
            st.subheader("Suma kwartalna — Ukończone vs Nieukończone")

            if not pivot.empty:
                per_q = (
                    pivot.groupby(["rok", "kwartał"], as_index=False)[
                        ["Ukończone", "Nieukończone"]
                    ]
                    .sum()
                    .sort_values(["rok", "kwartał"])
                )
                q_labels = [
                    f"{int(r)} | KW{int(k)}"
                    for r, k in zip(per_q["rok"], per_q["kwartał"])
                ]

                fig2 = go.Figure()
                fig2.add_trace(
                    go.Bar(
                        x=q_labels,
                        y=per_q["Ukończone"],
                        name="Ukończone",
                        marker_color="#43A047",
                        text=per_q["Ukończone"],
                        textposition="auto",
                    )
                )
                fig2.add_trace(
                    go.Bar(
                        x=q_labels,
                        y=per_q["Nieukończone"],
                        name="Nieukończone",
                        marker_color="#E53935",
                        text=per_q["Nieukończone"],
                        textposition="auto",
                    )
                )
                fig2.update_layout(
                    barmode="group",
                    template="plotly_dark",
                    paper_bgcolor="#0f0f1a",
                    plot_bgcolor="#1a1a2e",
                    font=dict(family="Inter", color="#e0e0e0"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    height=400,
                )
                st.plotly_chart(fig2, use_container_width=True)

            # ── Chart 3: Per-person completion rate ──────────────────────
            st.subheader("Completion Rate per Person")

            if not pivot.empty:
                per_person = (
                    pivot.groupby("osoba", as_index=False)[
                        ["Ukończone", "Nieukończone"]
                    ].sum()
                )
                per_person["total"] = per_person["Ukończone"] + per_person["Nieukończone"]
                per_person["rate"] = (
                    per_person["Ukończone"] / per_person["total"].replace(0, 1) * 100
                )
                per_person = per_person.sort_values("rate", ascending=True)

                fig3 = go.Figure()
                fig3.add_trace(
                    go.Bar(
                        y=per_person["osoba"],
                        x=per_person["rate"],
                        orientation="h",
                        marker_color=[
                            "#43A047" if r >= 70 else "#FFA726" if r >= 40 else "#E53935"
                            for r in per_person["rate"]
                        ],
                        text=[f"{r:.0f}%" for r in per_person["rate"]],
                        textposition="auto",
                    )
                )
                fig3.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="#0f0f1a",
                    plot_bgcolor="#1a1a2e",
                    font=dict(family="Inter", color="#e0e0e0"),
                    height=max(300, len(per_person) * 40),
                    xaxis=dict(title="Completion %", range=[0, 105]),
                    showlegend=False,
                )
                st.plotly_chart(fig3, use_container_width=True)

            # ── Data tables ──────────────────────────────────────────────
            st.divider()
            tab_agg, tab_det, tab_piv = st.tabs(["Agregat", "Szczegóły", "Pivot"])
            with tab_agg:
                st.dataframe(result.aggregate, use_container_width=True, height=400)
            with tab_det:
                st.dataframe(result.details, use_container_width=True, height=400)
            with tab_piv:
                st.dataframe(result.pivot, use_container_width=True, height=400)

            # ── AI Analysis ──────────────────────────────────────────────
            st.divider()
            st.subheader("🤖 AI Analysis")

            from src.ai.analyst import Analyst

            if not Analyst.is_available():
                st.warning(
                    "LiteLLM not installed. Run `uv add litellm` to enable AI analysis."
                )
            else:
                ai_col1, ai_col2 = st.columns([3, 1])
                with ai_col2:
                    llm_model = st.text_input("LLM Model", value="ollama/llama3.1:8b")
                    llm_url = st.text_input("LLM Base URL", value="http://localhost:11434")
                with ai_col1:
                    if st.button("🧠 Analyze with AI", use_container_width=True):
                        with st.spinner("Querying LLM..."):
                            analyst = Analyst(model=llm_model, base_url=llm_url)
                            ai_result = analyst.analyze(result.aggregate, result.pivot)
                            if ai_result.success:
                                st.markdown(f"**Summary:** {ai_result.summary}")
                                if ai_result.recommendations:
                                    st.markdown("**Recommendations:**")
                                    for rec in ai_result.recommendations:
                                        st.markdown(f"- {rec}")
                            else:
                                st.error(f"AI analysis failed: {ai_result.error}")
                                st.info(ai_result.summary)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — EXPORT
# ══════════════════════════════════════════════════════════════════════════════

with tab_export:
    st.header("Export Reports")

    result: EngineResult | None = st.session_state.result

    if result is None:
        st.info("Run analysis in the **Analytics** tab first to generate reports.")
    else:
        st.subheader("Excel Report")
        st.markdown(
            "Generates a multi-sheet Excel workbook with Pivot, Charts, Aggregate, and Details sheets. "
            "The chart uses a **hierarchical X-axis** (Osoba > Rok > KW) via xlsxwriter multi-level categories."
        )

        report_title = st.text_input(
            "Report Title",
            value="Ukończone vs Nieukończone (Osoba / Rok / Kwartał)",
        )

        if st.button("📥 Generate Excel Report", type="primary", use_container_width=True):
            with st.spinner("Generating Excel..."):
                xlsx_bytes = ExcelReportWriter.write_to_buffer(
                    pivot_df=result.pivot,
                    aggr_df=result.aggregate,
                    details_df=result.details,
                    title=report_title,
                )
                st.download_button(
                    label="⬇️ Download Excel Report",
                    data=xlsx_bytes,
                    file_name="planner_report_v10.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success("Report generated successfully!")

        st.divider()

        # ── CSV exports ──────────────────────────────────────────────────
        st.subheader("CSV / JSON Exports")
        c1, c2, c3 = st.columns(3)

        with c1:
            csv_agg = result.aggregate.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📄 Aggregate CSV",
                csv_agg,
                "aggregate.csv",
                "text/csv",
                use_container_width=True,
            )
        with c2:
            csv_det = result.details.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📄 Details CSV",
                csv_det,
                "details.csv",
                "text/csv",
                use_container_width=True,
            )
        with c3:
            json_pivot = result.pivot.to_json(orient="records", force_ascii=False, indent=2)
            st.download_button(
                "📄 Pivot JSON",
                json_pivot,
                "pivot.json",
                "application/json",
                use_container_width=True,
            )

        st.divider()
        st.subheader("Summary")
        st.json(
            {
                "source_rows": result.source_row_count,
                "filtered_rows": result.filtered_row_count,
                "pivot_rows": len(result.pivot),
                "aggregate_rows": len(result.aggregate),
                "details_rows": len(result.details),
                "unique_persons": int(result.pivot["osoba"].nunique()) if len(result.pivot) > 0 else 0,
            }
        )
