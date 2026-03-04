import pandas as pd
import pdfplumber
import os
import re
import sys
import glob
import logging
from datetime import datetime

from ai_matching import SemanticMatcher  # , assess_risk_batch  # ── LLM disabled

log = logging.getLogger(__name__)

if getattr(sys, 'frozen', False):
    COMPARE_FOLDER = os.path.dirname(sys.executable)
else:
    COMPARE_FOLDER = os.path.dirname(os.path.abspath(__file__))

APPROVED_PDFS = [
    os.path.join(COMPARE_FOLDER, "LIST OF SOFTWARES - (APPROVED BY SAM).pdf"),
    os.path.join(COMPARE_FOLDER, "LIST OF SOFTWARES(APPROVED BY SAM).pdf"),
]
NOT_APPROVED_PDFS = [
    os.path.join(COMPARE_FOLDER, "LIST OF SOFTWARES (NOT APPROVED BY SAM).pdf"),
    os.path.join(COMPARE_FOLDER, "LIST OF SOFTWARES (NOT APPROVED BY SAM) 26 SEPT 2025.pdf"),
]

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
    approved_dfs = [extract_pdf_table(f) for f in APPROVED_PDFS if os.path.isfile(f)]
    not_approved_dfs = [extract_pdf_table(f) for f in NOT_APPROVED_PDFS if os.path.isfile(f)]

    allowed = pd.concat(approved_dfs, ignore_index=True) if approved_dfs else pd.DataFrame(columns=["Software"])
    not_allowed = pd.concat(not_approved_dfs, ignore_index=True) if not_approved_dfs else pd.DataFrame(columns=["Software"])

    allowed["Software"] = allowed["Software"].str.lower().str.strip()
    not_allowed["Software"] = not_allowed["Software"].str.lower().str.strip()
    allowed["Status"] = "Allowed"
    not_allowed["Status"] = "Not Allowed"

    official = pd.concat([allowed, not_allowed], ignore_index=True).drop_duplicates("Software")
    return official


def clean_name(name):
    if pd.isna(name):
        return ""
    name = str(name).lower()
    name = re.sub(r"\d+\.\d+[\.\d]*", "", name)
    name = re.sub(r"\(x64\)|\(64-bit\)|\(32-bit\)", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[^\w\s]", " ", name)
    return " ".join(name.split()).strip()


def read_uploaded_csv(csv_path):
    rows = []
    meta = {}
    with open(csv_path, "r", encoding="utf-8") as f:
        for line in f:
            if line.startswith("#"):
                parts = line.lstrip("#").strip().split(":", 1)
                if len(parts) == 2:
                    meta[parts[0].strip()] = parts[1].strip()
            else:
                rows.append(line)

    if not rows:
        print("  ERROR: CSV file is empty.")
        sys.exit(1)

    from io import StringIO
    df = pd.read_csv(StringIO("".join(rows)))
    return df, meta


def process_upload(csv_path):
    user_df, meta = read_uploaded_csv(csv_path)
    hostname = meta.get('Hostname', '?')
    print(f"  Checking {hostname} ({len(user_df)} programs)...")

    official = build_official_list()

    # ── AI semantic matcher ──
    semantic = SemanticMatcher(official, threshold=0.55)
    if semantic.available:
        print("  AI semantic matching: ACTIVE")
    else:
        print("  AI semantic matching: unavailable (using fuzzy fallback)")

    results = []
    counts = {"Allowed": 0, "Not Allowed": 0, "Not Found": 0}

    official_clean = official["Software"].apply(clean_name).tolist()
    official_status = official["Status"].tolist()
    official_names = official["Software"].tolist()

    # Pass 1: exact/substring match
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
                break

        results.append({
            "Installed_Software": sw_name,
            "Version": str(row.get("Version") or ""),
            "Publisher": str(row.get("Publisher") or ""),
            "Status": status,
            "Match_Type": match_type,
            "Matched_Official": matched,
            "Confidence_%": confidence,
            "AI_Risk_Level": "",
            "AI_Risk_Reasoning": "",
        })

    # Pass 2: batch semantic match for all unmatched items in one call
    unmatched_idx = [i for i, r in enumerate(results) if r["Match_Type"] == "—"]
    if unmatched_idx and semantic.available:
        unmatched_names = [results[i]["Installed_Software"] for i in unmatched_idx]
        sem_results = semantic.find_best_matches_batch(unmatched_names)
        for i, (sem_matched, sem_conf, sem_status) in zip(unmatched_idx, sem_results):
            if sem_matched:
                results[i]["Matched_Official"] = sem_matched
                results[i]["Confidence_%"] = sem_conf
                results[i]["Status"] = sem_status
                results[i]["Match_Type"] = "AI Semantic"

    for r in results:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1

    # ── LLM risk assessment disabled (too slow) ──
    # not_found = [r for r in results if r["Status"] == "Not Found"]
    # if not_found:
    #     risks = assess_risk_batch(not_found)
    #     for r in results:
    #         risk = risks.get(r["Installed_Software"])
    #         if risk:
    #             r["AI_Risk_Level"] = risk["risk_level"]
    #             r["AI_Risk_Reasoning"] = risk["reasoning"]

    result_df = pd.DataFrame(results)
    result_df.sort_values(["Status", "Installed_Software"], inplace=True)

    total = len(result_df)
    print()
    for status in ["Allowed", "Not Allowed", "Not Found"]:
        cnt = counts.get(status, 0)
        pct = cnt / total * 100 if total else 0
        print(f"   {status:.<25} {cnt:>4}  ({pct:.1f}%)")
    print(f"   {'Total':.<25} {total:>4}")

    base = os.path.splitext(os.path.basename(csv_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_name = f"compliance_report_{timestamp}.xlsx"
    report_path = os.path.join(COMPARE_FOLDER, report_name)
    result_df.to_excel(report_path, index=False)
    print(f"\n  Report saved: {report_name}")

    not_allowed = result_df[result_df["Status"] == "Not Allowed"]
    if not not_allowed.empty:
        print(f"\n  NOT ALLOWED ({len(not_allowed)}):")
        for _, r in not_allowed.iterrows():
            print(f"    - {r['Installed_Software']}")

    not_found = result_df[result_df["Status"] == "Not Found"]
    if not not_found.empty:
        print(f"\n  UNKNOWN ({len(not_found)}) - check report:")
        for _, r in not_found.head(10).iterrows():
            print(f"    - {r['Installed_Software']}")
        if len(not_found) > 10:
            print(f"    ... and {len(not_found) - 10} more")

    return report_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        pattern = os.path.join(COMPARE_FOLDER, "installed_software_*.csv")
        csvs = sorted(glob.glob(pattern))
        if csvs:
            csv_file = csvs[-1]
        else:
            print("  No scan files found.")
            sys.exit(1)
    else:
        csv_file = sys.argv[1]

    if not os.path.isfile(csv_file):
        print(f"  File not found: {csv_file}")
        sys.exit(1)

    process_upload(csv_file)
