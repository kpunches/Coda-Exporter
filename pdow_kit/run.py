#!/usr/bin/env python3
"""
Program Map kit entry point.

Usage:
    python run.py                    # uses ./data.json, writes ./output/<PROG>_Program_Map.xlsx
    python run.py path/to/data.json  # explicit input
    python run.py --out custom.xlsx  # custom output path

Reads a bundled `data.json` (produced by the UI bundler), merges alignments,
and emits the formatted xlsx using the same writer the in-repo
pdow_exporter project uses.
"""
from __future__ import annotations
import argparse
import json
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from merger import merge
from xlsx_writer import write_program_map


def main():
    parser = argparse.ArgumentParser(description="Build a Program Map xlsx from a bundled data.json")
    parser.add_argument("data_json", nargs="?", default=str(HERE / "data.json"),
                        help="Path to data.json (default: ./data.json)")
    parser.add_argument("--out", default=None,
                        help="Output xlsx path (default: ./output/<PROG>_Program_Map.xlsx)")
    args = parser.parse_args()

    data_path = Path(args.data_json)
    if not data_path.exists():
        print(f"ERROR: {data_path} does not exist", file=sys.stderr)
        return 1

    t0 = time.time()
    data = json.loads(data_path.read_text())
    program_code = data.get("program_code", "PROGRAM")

    model = merge(data)

    out_path = Path(args.out) if args.out else (HERE / "output" / f"{program_code}_Program_Map.xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    write_program_map(
        model=model,
        template_path=HERE / "template_BSSCOM.xlsx",
        output_path=out_path,
        owl_image_path=HERE / "owl.png",
        tab_name="Program Map",
        program_code=program_code,
    )
    print(f"[XLSX] Wrote {out_path} in {time.time() - t0:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
