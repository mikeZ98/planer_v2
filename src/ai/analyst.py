"""
AI Analyst v10.0 — LLM-powered semantic analysis of engine outputs.

Connects to a local LLM via LiteLLM (supports Ollama, OpenAI-compatible APIs,
and any provider LiteLLM supports). Generates natural-language insights from
aggregated task data.

Usage:
    analyst = Analyst(base_url="http://localhost:11434")
    insight = analyst.analyze(engine_result)
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Any, Optional

import pandas as pd


@dataclass
class AnalysisResult:
    """Container for LLM analysis output."""
    summary: str
    recommendations: list[str]
    raw_response: str
    model: str
    success: bool
    error: Optional[str] = None


# ── Prompt engineering ───────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
You are a senior project management analyst. You receive structured task data
aggregated by Person, Year, and Quarter. Each record shows completed vs incomplete
tasks. Your job is to:

1. Identify bottlenecks — people or periods with high incomplete task ratios.
2. Spot trends — improving or worsening completion rates over quarters.
3. Give actionable recommendations — concrete, data-driven suggestions.

Respond in Polish. Be concise and direct. Use bullet points.
Format your response as JSON with keys: "summary" (string), "recommendations" (list of strings).
"""


def _build_user_prompt(
    aggregate_df: pd.DataFrame,
    pivot_df: pd.DataFrame,
) -> str:
    """Build the data context prompt from engine DataFrames."""
    agg_text = aggregate_df.to_string(index=False, max_rows=100)
    pivot_text = pivot_df.to_string(index=False, max_rows=100)

    total_done = int(aggregate_df.get("zadania_ukończone", pd.Series([0])).sum())
    total_undone = int(aggregate_df.get("zadania_nieukończone", pd.Series([0])).sum())

    return f"""\
Dane zagregowane (Zasobnik × Rok × Kwartał):
{agg_text}

Dane pivot (Osoba × Rok × Kwartał):
{pivot_text}

Podsumowanie: {total_done} ukończonych, {total_undone} nieukończonych zadań.
Współczynnik ukończenia: {total_done / max(total_done + total_undone, 1) * 100:.1f}%

Przeanalizuj dane i podaj wnioski.
"""


# ── Analyst ──────────────────────────────────────────────────────────────────

class Analyst:
    """
    LLM-powered analyst using LiteLLM for provider-agnostic inference.

    Supports:
      - Ollama (local): base_url="http://localhost:11434"
      - OpenAI-compatible: set OPENAI_API_KEY env var
      - Any LiteLLM-supported provider
    """

    DEFAULT_MODEL = "ollama/llama3.1:8b"

    def __init__(
        self,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
        api_key: Optional[str] = None,
    ):
        self.model = model or os.getenv("PLANNER_LLM_MODEL", self.DEFAULT_MODEL)
        self.base_url = base_url or os.getenv("PLANNER_LLM_BASE_URL", "http://localhost:11434")
        self.api_key = api_key or os.getenv("PLANNER_LLM_API_KEY", "")

    def analyze(
        self,
        aggregate_df: pd.DataFrame,
        pivot_df: pd.DataFrame,
        timeout: float = 60.0,
    ) -> AnalysisResult:
        """
        Send engine outputs to the LLM and return structured insights.
        Gracefully degrades if LLM is unavailable.
        """
        try:
            import litellm

            if self.base_url:
                litellm.api_base = self.base_url

            user_prompt = _build_user_prompt(aggregate_df, pivot_df)

            response = litellm.completion(
                model=self.model,
                messages=[
                    {"role": "system", "content": _SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                api_key=self.api_key or None,
                api_base=self.base_url,
                timeout=timeout,
                temperature=0.3,
                max_tokens=2000,
            )

            raw = response.choices[0].message.content
            return self._parse_response(raw, self.model)

        except ImportError:
            return AnalysisResult(
                summary="Moduł litellm nie jest zainstalowany. Zainstaluj: uv add litellm",
                recommendations=["pip install litellm"],
                raw_response="",
                model=self.model,
                success=False,
                error="litellm not installed",
            )
        except Exception as e:
            return AnalysisResult(
                summary=f"Nie udało się połączyć z LLM: {e}",
                recommendations=[
                    "Sprawdź czy Ollama jest uruchomiony (ollama serve)",
                    f"Sprawdź model: {self.model}",
                    f"Sprawdź URL: {self.base_url}",
                ],
                raw_response="",
                model=self.model,
                success=False,
                error=str(e),
            )

    @staticmethod
    def _parse_response(raw: str, model: str) -> AnalysisResult:
        """Try to parse JSON response; fall back to raw text."""
        try:
            # Try to extract JSON from response (LLM may wrap in markdown)
            json_str = raw
            if "```json" in raw:
                json_str = raw.split("```json")[1].split("```")[0]
            elif "```" in raw:
                json_str = raw.split("```")[1].split("```")[0]

            data = json.loads(json_str.strip())
            return AnalysisResult(
                summary=data.get("summary", raw),
                recommendations=data.get("recommendations", []),
                raw_response=raw,
                model=model,
                success=True,
            )
        except (json.JSONDecodeError, IndexError):
            return AnalysisResult(
                summary=raw,
                recommendations=[],
                raw_response=raw,
                model=model,
                success=True,
            )

    @staticmethod
    def is_available() -> bool:
        """Check if litellm is importable."""
        try:
            import litellm  # noqa: F401
            return True
        except ImportError:
            return False
