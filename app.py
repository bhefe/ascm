import os
import re
import io
import base64
import logging
import threading
from contextlib import asynccontextmanager
from datetime import datetime

import pandas as pd
import pdfplumber
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from ai_matching import SemanticMatcher  # , assess_risk_batch  # ── LLM disabled

log = logging.getLogger(__name__)

# ── Pre-load at startup (runs once) ──────────────────────────────
_official_list = None
_semantic_matcher = None

def get_official_list():
    global _official_list
    if _official_list is None:
        _official_list = build_official_list()
    return _official_list

def get_semantic_matcher():
    global _semantic_matcher
    if _semantic_matcher is None:
        print("[AI] Loading sentence-transformer model and encoding official list...", flush=True)
        _semantic_matcher = SemanticMatcher(get_official_list(), threshold=0.55)
        if _semantic_matcher.available:
            print(f"[AI] Semantic matcher ready — {len(_semantic_matcher.names)} software names encoded.", flush=True)
        else:
            print("[AI] Semantic matcher unavailable, will use fuzzy fallback.", flush=True)
    return _semantic_matcher

@asynccontextmanager
async def lifespan(app):
    """Pre-load PDFs and embedding model at startup so requests are instant."""
    def _load():
        print("[STARTUP] Pre-loading PDFs and AI model — this runs once...", flush=True)
        get_semantic_matcher()
        print("[STARTUP] Ready. Uploads will now be processed quickly.", flush=True)
    threading.Thread(target=_load, daemon=True).start()
    yield

app = FastAPI(title="Software Compliance API", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APPROVED_PDFS = [
    os.path.join(BASE_DIR, "LIST OF SOFTWARES - (APPROVED BY SAM).pdf"),
    os.path.join(BASE_DIR, "LIST OF SOFTWARES(APPROVED BY SAM).pdf"),
]
NOT_APPROVED_PDFS = [
    os.path.join(BASE_DIR, "LIST OF SOFTWARES (NOT APPROVED BY SAM).pdf"),
    os.path.join(BASE_DIR, "LIST OF SOFTWARES (NOT APPROVED BY SAM) 26 SEPT 2025.pdf"),
]

# ── Compliance logic ──────────────────────────────────────

def extract_pdf_table(pdf_path):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                data.extend(table)
    if not data:
        return pd.DataFrame(columns=["Software"])
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = [c.strip() if c else "" for c in df.columns]
    sw_col = [c for c in df.columns if "software" in c.lower()]
    if sw_col:
        df.rename(columns={sw_col[0]: "Software"}, inplace=True)
    elif len(df.columns) > 1:
        df.rename(columns={df.columns[1]: "Software"}, inplace=True)
    return df[["Software"]].dropna()


def build_official_list():
    print("[PDF] Parsing approved software list from PDFs...", flush=True)
    approved_dfs = [extract_pdf_table(f) for f in APPROVED_PDFS if os.path.isfile(f)]
    not_approved_dfs = [extract_pdf_table(f) for f in NOT_APPROVED_PDFS if os.path.isfile(f)]
    allowed = pd.concat(approved_dfs, ignore_index=True) if approved_dfs else pd.DataFrame(columns=["Software"])
    not_allowed = pd.concat(not_approved_dfs, ignore_index=True) if not_approved_dfs else pd.DataFrame(columns=["Software"])
    allowed["Software"] = allowed["Software"].str.lower().str.strip()
    not_allowed["Software"] = not_allowed["Software"].str.lower().str.strip()
    allowed["Status"] = "Allowed"
    not_allowed["Status"] = "Not Allowed"
    result = pd.concat([allowed, not_allowed], ignore_index=True).drop_duplicates("Software")
    print(f"[PDF] Done — {len(allowed)} allowed, {len(not_allowed)} not allowed ({len(result)} unique entries)", flush=True)
    return result


def clean_name(name):
    if pd.isna(name):
        return ""
    name = str(name).lower()
    name = re.sub(r"\d+\.\d+[\.\d]*", "", name)
    name = re.sub(r"\(x64\)|\(64-bit\)|\(32-bit\)", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[^\w\s]", " ", name)
    return " ".join(name.split()).strip()


def read_csv(file_bytes):
    rows, meta = [], {}
    for line in file_bytes.decode("utf-8").splitlines(keepends=True):
        if line.startswith("#"):
            parts = line.lstrip("#").strip().split(":", 1)
            if len(parts) == 2:
                meta[parts[0].strip()] = parts[1].strip()
        else:
            rows.append(line)
    df = pd.read_csv(io.StringIO("".join(rows)))
    return df, meta


def run_check(file_bytes):
    user_df, meta = read_csv(file_bytes)
    hostname = meta.get('Hostname', 'unknown')
    print(f"\n[CHECK] Request from: {hostname} — {len(user_df)} programs to check", flush=True)

    official = get_official_list()       # cached
    semantic = get_semantic_matcher()    # cached

    results = []
    counts = {"Allowed": 0, "Not Allowed": 0, "Not Found": 0}
    ai_count = 0
    exact_count = 0

    # Pre-compute cleaned official names once (avoids re-cleaning on every comparison)
    official_clean = official["Software"].apply(clean_name).tolist()
    official_status = official["Status"].tolist()
    official_names = official["Software"].tolist()

    print(f"[CHECK] Matching software against {len(official)} official entries...", flush=True)

    # Pass 1: exact/substring match for every row
    for _, row in user_df.iterrows():
        sw_name = row["Software"]
        sw_clean = clean_name(sw_name)
        matched = ""
        status = "Not Found"
        match_type = "—"
        confidence = 0

        for idx, off_clean in enumerate(official_clean):
            if sw_clean in off_clean or off_clean in sw_clean:
                matched = official_names[idx]
                status = official_status[idx]
                match_type = "Exact"
                confidence = 100
                exact_count += 1
                break

        results.append({
            "software": sw_name,
            "version": str(row.get("Version") or ""),
            "publisher": str(row.get("Publisher") or ""),
            "status": status,
            "matched": matched,
            "confidence": confidence,
            "match_type": match_type,
            "risk_level": "",
            "risk_reasoning": "",
        })

    # Pass 2: batch semantic match for all unmatched items in one call
    unmatched_idx = [i for i, r in enumerate(results) if r["match_type"] == "—"]
    if unmatched_idx and semantic.available:
        unmatched_names = [results[i]["software"] for i in unmatched_idx]
        print(f"[CHECK] Running batch semantic match on {len(unmatched_names)} unmatched items...", flush=True)
        sem_results = semantic.find_best_matches_batch(unmatched_names)
        for i, (sem_matched, sem_conf, sem_status) in zip(unmatched_idx, sem_results):
            if sem_matched:
                results[i]["matched"] = sem_matched
                results[i]["confidence"] = sem_conf
                results[i]["status"] = sem_status
                results[i]["match_type"] = "AI Semantic"
                ai_count += 1

    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    print(f"[CHECK] Done — Exact: {exact_count}, AI Semantic: {ai_count}, Not Found: {counts.get('Not Found', 0)}", flush=True)
    print(f"[CHECK] Summary — Allowed: {counts['Allowed']}, Not Allowed: {counts['Not Allowed']}, Not Found: {counts['Not Found']}", flush=True)

    # ── LLM risk assessment disabled (too slow) ──
    # not_found = [r for r in results if r["status"] == "Not Found"]
    # if not_found:
    #     risks = assess_risk_batch(not_found)
    #     for r in results:
    #         risk = risks.get(r["software"])
    #         if risk:
    #             r["risk_level"] = risk["risk_level"]
    #             r["risk_reasoning"] = risk["reasoning"]

    results.sort(key=lambda r: (["Not Allowed", "Not Found", "Allowed"].index(r["status"]), r["software"].lower()))

    return results, counts, meta


def build_excel(results):
    df = pd.DataFrame([{
        "Installed Software": r["software"],
        "Version": r["version"],
        "Publisher": r["publisher"],
        "Status": r["status"],
        "Match Type": r.get("match_type", ""),
        "Matched Official": r["matched"],
        "Confidence %": r["confidence"],
        "AI Risk Level": r.get("risk_level", ""),
        "AI Risk Reasoning": r.get("risk_reasoning", ""),
    } for r in results])
    df.sort_values(["Status", "Installed Software"], inplace=True)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ── API routes ────────────────────────────────────────────

@app.post("/api/check")
async def check(csv_file: UploadFile = File(...)):
    if not csv_file.filename.endswith(".csv"):
        raise HTTPException(status_code=400, detail="Please upload a valid .csv file.")
    file_bytes = await csv_file.read()
    try:
        results, counts, meta = run_check(file_bytes)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error processing file: {e}")

    csv_b64 = base64.b64encode(file_bytes).decode("utf-8")
    total = sum(counts.values())
    return {
        "results": results,
        "counts": counts,
        "meta": meta,
        "total": total,
        "filename": csv_file.filename,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "csv_b64": csv_b64,
    }


@app.post("/api/download")
async def download(csv_b64: str = Form(...), filename: str = Form("report")):
    try:
        file_bytes = base64.b64decode(csv_b64)
        results, _, _ = run_check(file_bytes)
        buf = build_excel(results)
        report_name = f"compliance_{filename.replace('.csv', '')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{report_name}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating report: {e}")


# ── Serve Angular build ───────────────────────────────────

ANGULAR_DIST = os.path.join(BASE_DIR, "frontend", "dist", "frontend", "browser")
if os.path.isdir(ANGULAR_DIST):
    from fastapi.responses import FileResponse

    assets_dir = os.path.join(ANGULAR_DIST, "assets")
    if os.path.isdir(assets_dir):
        app.mount("/assets", StaticFiles(directory=assets_dir), name="assets")

    app.mount("/static", StaticFiles(directory=ANGULAR_DIST), name="static")

    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str):
        file_path = os.path.join(ANGULAR_DIST, full_path)
        if os.path.isfile(file_path):
            return FileResponse(file_path)
        return FileResponse(os.path.join(ANGULAR_DIST, "index.html"))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
