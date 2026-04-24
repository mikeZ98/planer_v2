"""
Planner v10.0 — Entry point.

Usage:
    uv run streamlit run app.py                    # Streamlit dashboard (default)
    uv run uvicorn src.api.service:app --port 8000 # FastAPI service
    uv run python main.py                          # Quick engine smoke test
"""

from src.engine.processor import PlannerEngine, ColumnMap
from src.engine.ingestor import UniversalIngestor


def main():
    print("Planner v10.0 — Universal Data Engineering Service")
    print("─" * 50)
    print()
    print("Launch options:")
    print("  Streamlit UI : uv run streamlit run app.py")
    print("  FastAPI       : uv run uvicorn src.api.service:app --reload --port 8000")
    print("  Tests         : uv run pytest tests/ -v")
    print()

    # Quick smoke test: check that engine imports work
    try:
        engine = PlannerEngine()
        print("✓ Engine loaded successfully")
        print(f"✓ Ingestor supports: {', '.join(UniversalIngestor.EXCEL_EXTENSIONS | UniversalIngestor.CSV_EXTENSIONS | UniversalIngestor.JSON_EXTENSIONS | UniversalIngestor.TEXT_EXTENSIONS)}")
        print("✓ All systems nominal.")
    except Exception as e:
        print(f"✗ Engine load failed: {e}")


if __name__ == "__main__":
    main()
