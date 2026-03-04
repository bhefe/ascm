"""
AI-powered software matching and risk assessment.

1. SemanticMatcher  – uses sentence-transformers embeddings + cosine
   similarity for far more accurate software-name matching.
2. RiskAssessor     – calls a local Ollama LLM to assess risk for
   software not found in the official SAM lists and generates a
   pre-filled approval request.

Both components degrade gracefully: if the model isn't downloaded yet
or Ollama isn't running, the caller gets a clear fallback signal.
"""

import os
import re
import logging

# Workaround for duplicate OpenMP runtime in Anaconda environments
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")

import numpy as np
import pandas as pd
import httpx

log = logging.getLogger(__name__)

# ── Semantic Matcher ──────────────────────────────────────────────

# Model is loaded lazily on first use so the import itself is instant.
_model = None
EMBED_MODEL_NAME = "all-MiniLM-L6-v2"          # ~80 MB, very fast


def _get_model():
    """Load (and cache) the sentence-transformer model."""
    global _model
    if _model is None:
        try:
            from sentence_transformers import SentenceTransformer
            log.info("Loading embedding model '%s' …", EMBED_MODEL_NAME)
            _model = SentenceTransformer(EMBED_MODEL_NAME)
            log.info("Embedding model ready.")
        except Exception as exc:
            log.warning("Could not load embedding model: %s", exc)
            _model = False          # sentinel: tried and failed
    return _model if _model else None


def _clean(name: str) -> str:
    """Normalise a software name for embedding."""
    if pd.isna(name):
        return ""
    name = str(name).lower()
    name = re.sub(r"\d+\.\d+[\.\d]*", "", name)
    name = re.sub(r"\(x64\)|\(64-bit\)|\(32-bit\)", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[^\w\s]", " ", name)
    return " ".join(name.split()).strip()


class SemanticMatcher:
    """
    Pre-encodes the official SAM software list into embeddings, then
    matches each scanned software name by cosine similarity.
    """

    def __init__(self, official_df: pd.DataFrame, threshold: float = 0.55):
        """
        Parameters
        ----------
        official_df : DataFrame with columns ``Software`` and ``Status``.
        threshold   : minimum cosine similarity to consider a match.
        """
        self.threshold = threshold
        self.available = False

        model = _get_model()
        if model is None:
            log.warning("SemanticMatcher unavailable – will fall back to fuzzy.")
            self.names = []
            self.statuses = []
            self.embeddings = None
            return

        # Store the raw (lowered) names and statuses
        self.names = official_df["Software"].tolist()
        self.statuses = official_df["Status"].tolist()

        # Build cleaned representations for encoding
        cleaned = [_clean(n) for n in self.names]
        self.embeddings = model.encode(cleaned, normalize_embeddings=True,
                                       show_progress_bar=False)
        self.available = True
        log.info("Encoded %d official software names.", len(self.names))

    # ------------------------------------------------------------------ #

    def find_best_match(self, user_sw: str):
        """
        Returns (matched_name, confidence_0_to_100, status) or
        (None, 0, None) when nothing passes the threshold.
        """
        if not self.available:
            return None, 0, None

        model = _get_model()
        user_emb = model.encode([_clean(user_sw)], normalize_embeddings=True,
                                show_progress_bar=False)

        # cosine similarity (embeddings are already L2-normalised)
        scores = (self.embeddings @ user_emb.T).flatten()
        best_idx = int(np.argmax(scores))
        best_score = float(scores[best_idx])

        if best_score >= self.threshold:
            return (self.names[best_idx],
                    round(best_score * 100, 1),
                    self.statuses[best_idx])
        return None, 0, None

    def find_best_matches_batch(self, names: list) -> list:
        """
        Encode all names in one shot and return a list of
        (matched_name, confidence, status) tuples — one per input name.
        Much faster than calling find_best_match() in a loop.
        """
        if not self.available or not names:
            return [(None, 0, None)] * len(names)

        model = _get_model()
        cleaned = [_clean(n) for n in names]
        user_embs = model.encode(cleaned, normalize_embeddings=True,
                                 show_progress_bar=False, batch_size=64)

        # Single matrix multiply: (N_user, D) @ (D, N_official) = (N_user, N_official)
        scores_matrix = user_embs @ self.embeddings.T

        results = []
        for scores in scores_matrix:
            best_idx = int(np.argmax(scores))
            best_score = float(scores[best_idx])
            if best_score >= self.threshold:
                results.append((self.names[best_idx],
                                round(best_score * 100, 1),
                                self.statuses[best_idx]))
            else:
                results.append((None, 0, None))
        return results


# ── LLM Risk Assessor (Ollama) ────────────────────────────────────

OLLAMA_URL = "http://localhost:11434"
OLLAMA_MODEL = "gemma3:1b"             # using locally installed model
OLLAMA_TIMEOUT = 60                # seconds per request


def _ollama_available() -> bool:
    """Quick health-check for a running Ollama instance."""
    try:
        r = httpx.get(OLLAMA_URL, timeout=3)
        return r.status_code == 200
    except Exception:
        return False


# ── Batch helper (used by both app.py and check_upload.py) ────────

def assess_risk_batch(not_found_items: list[dict]) -> dict[str, dict]:
    """
    Assess risk for ALL unknown software in a single Ollama call.
    Returns a mapping of software-name → risk-dict.
    """
    if not not_found_items or not _ollama_available():
        return {}

    # Build numbered list for the prompt
    items = []
    for i, item in enumerate(not_found_items, 1):
        name = item.get("software") or item.get("Installed_Software", "")
        publisher = item.get("publisher") or item.get("Publisher", "") or "Unknown"
        version = item.get("version") or item.get("Version", "") or "Unknown"
        items.append((i, name, publisher, version))

    software_list = "\n".join(
        f"{i}. {name} | Publisher: {pub} | Version: {ver}"
        for i, name, pub, ver in items
    )

    prompt = f"""You are an IT security analyst evaluating software compliance.
The following software items are installed on a user's workstation but are NOT in our approved or blocked software lists.

{software_list}

For EACH item respond in EXACTLY this repeating format (one block per item, no extra text):

ITEM: <number>
RISK_LEVEL: <Low|Medium|High>
REASONING: <one sentence>
---
"""

    print(f"[AI]   Sending {len(items)} items to Ollama in one request...", flush=True)
    try:
        resp = httpx.post(
            f"{OLLAMA_URL}/api/generate",
            json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=max(OLLAMA_TIMEOUT, len(items) * 5),
        )
        resp.raise_for_status()
        raw = resp.json().get("response", "")
        return _parse_batch_response(raw, items)
    except Exception as exc:
        log.warning("Ollama batch risk assessment failed: %s", exc)
        return {}


def _parse_batch_response(text: str, items: list) -> dict[str, dict]:
    """Parse a batch LLM response into a name → risk dict."""
    # Index items by number
    num_to_name = {i: name for i, name, _, _ in items}
    results = {}

    current_num = None
    current = {"risk_level": "Unknown", "reasoning": ""}

    for line in text.strip().splitlines():
        line = line.strip()
        if not line or line == "---":
            if current_num and current_num in num_to_name:
                results[num_to_name[current_num]] = dict(current)
            current_num = None
            current = {"risk_level": "Unknown", "reasoning": ""}
            continue

        upper = line.upper()
        if upper.startswith("ITEM:"):
            # Save previous block if any
            if current_num and current_num in num_to_name:
                results[num_to_name[current_num]] = dict(current)
            try:
                current_num = int(line.split(":", 1)[1].strip())
            except ValueError:
                current_num = None
            current = {"risk_level": "Unknown", "reasoning": ""}
        elif upper.startswith("RISK_LEVEL:"):
            val = line.split(":", 1)[1].strip().capitalize()
            if val in ("Low", "Medium", "High"):
                current["risk_level"] = val
        elif upper.startswith("REASONING:"):
            current["reasoning"] = line.split(":", 1)[1].strip()

    # Save last block
    if current_num and current_num in num_to_name:
        results[num_to_name[current_num]] = dict(current)

    print(f"[AI]   Parsed {len(results)}/{len(items)} results from batch response.", flush=True)
    return results
