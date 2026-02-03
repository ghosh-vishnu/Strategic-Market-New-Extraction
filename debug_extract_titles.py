import argparse
import os
import re
import sys
from dataclasses import dataclass
from typing import Iterable


YEAR_RANGE_RE = re.compile(r"20\d{2}\s*[\-–]\s*20\d{2}")


@dataclass
class TitleResult:
    path: str
    title: str
    kind: str


def _classify(title: str) -> str:
    low = (title or "").lower()
    has_year_range = bool(YEAR_RANGE_RE.search(title or ""))
    if "segment revenue estimation" in low and "forecast" in low and has_year_range:
        return "segmented"
    if "forecast" in low and has_year_range:
        return "forecast"
    if "market" in low and has_year_range:
        return "market_only"
    return "other"


def _iter_docx_files(root_dir: str) -> Iterable[str]:
    for name in sorted(os.listdir(root_dir), key=lambda s: s.lower()):
        if name.lower().endswith(".docx"):
            yield os.path.join(root_dir, name)


def main() -> int:
    parser = argparse.ArgumentParser(description="Batch debug extract_title() over a directory of .docx files.")
    parser.add_argument(
        "--project-root",
        default=os.path.dirname(os.path.abspath(__file__)),
        help="Project root (used to import backend module).",
    )
    parser.add_argument(
        "--input-dir",
        default=None,
        help="Directory containing .docx files (default: <project-root>/Extracted).",
    )
    parser.add_argument("--show", type=int, default=30, help="How many non-segmented examples to print.")
    args = parser.parse_args()

    project_root = os.path.abspath(args.project_root)
    input_dir = os.path.abspath(args.input_dir or os.path.join(project_root, "Extracted"))

    sys.path.insert(0, os.path.join(project_root, "backend"))

    from converter.utils.extractor import extract_title  # noqa: PLC0415

    files = list(_iter_docx_files(input_dir))
    results: list[TitleResult] = []
    exceptions: list[tuple[str, str, str]] = []

    for path in files:
        try:
            title = extract_title(path)
        except Exception as e:  # pragma: no cover
            exceptions.append((path, type(e).__name__, str(e)))
            continue
        results.append(TitleResult(path=path, title=title, kind=_classify(title)))

    print(f"Input dir: {input_dir}")
    print(f"Total docx: {len(files)}")
    print(f"OK: {len(results)}")
    print(f"Exceptions: {len(exceptions)}")

    counts: dict[str, int] = {"segmented": 0, "forecast": 0, "market_only": 0, "other": 0}
    for r in results:
        counts[r.kind] = counts.get(r.kind, 0) + 1
    for k in ["segmented", "forecast", "market_only", "other"]:
        print(f"{k}: {counts.get(k, 0)}")

    if exceptions:
        print("\n--- Exceptions (first 20) ---")
        for path, typ, msg in exceptions[:20]:
            print(f"{os.path.basename(path)} => {typ}: {msg}")

    shown = 0
    if args.show > 0:
        print(f"\n--- Non-segmented examples (first {args.show}) ---")
        for k in ["forecast", "market_only", "other"]:
            for r in results:
                if r.kind != k:
                    continue
                print(f"{os.path.basename(r.path)} => {r.title}")
                shown += 1
                if shown >= args.show:
                    break
            if shown >= args.show:
                break

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

