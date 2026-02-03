"""Compare extract_title results: Okay (reference) vs Extracted."""
import os
import sys
import re

root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(root, "backend"))
from converter.utils.extractor import extract_title


def classify(t):
    if not t or t == "Title Not Available":
        return "fail"
    low = t.lower()
    has_seg = "segment revenue estimation" in low and "forecast" in low
    has_by = len(re.findall(r"\bby\s+\w", low)) >= 2
    has_yr = bool(re.search(r"20\d{2}\s*[\-–]\s*20\d{2}", t))
    if has_seg and has_yr:
        return "segmented"
    if has_by and has_yr:
        return "partial"
    if "market" in low:
        return "basic"
    return "other"


def run_folder(folder_name, max_files=120):
    path = os.path.join(root, folder_name)
    if not os.path.isdir(path):
        return None, []
    files = sorted(
        [f for f in os.listdir(path) if f.lower().endswith(".docx")],
        key=str.lower,
    )[:max_files]
    results = {}
    bad = []
    for f in files:
        p = os.path.join(path, f)
        try:
            t = extract_title(p)
        except Exception as e:
            results["error"] = results.get("error", 0) + 1
            bad.append((f, "ERROR: " + str(e)[:80]))
            continue
        k = classify(t)
        results[k] = results.get(k, 0) + 1
        if k in ("fail", "basic", "other"):
            bad.append((f, (t[:100] + "...") if len(t) > 100 else t))
    return results, bad


if __name__ == "__main__":
    for folder in ["Okay", "Extracted"]:
        res, bad = run_folder(folder)
        if res is None:
            print(folder, ": folder not found")
            continue
        print(folder, ":", res)
        for f, msg in bad[:12]:
            print("  BAD:", f, "=>", msg)
        print()
