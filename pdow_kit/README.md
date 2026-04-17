# PDOW Kit

Self-contained Python kit that reads a bundled `data.json` (program model + Coda alignments) and emits a fully formatted Program Map xlsx — same output as the in-repo `pdow_exporter` project, just driven by JSON instead of hardcoded Python tuples.

This kit is what the web UI bundles into a downloadable zip after fetching live data via the Coda MCP. End users unzip it and run `python run.py` locally; no Coda credentials needed on their machine because the data is baked into `data.json` at bundle time.

## Run

```bash
python run.py                    # uses ./data.json, writes ./output/<PROG>_Program_Map.xlsx
python run.py path/to/data.json  # explicit input
python run.py --out custom.xlsx  # custom output path
```

Requires Python 3.10+ with `openpyxl` and `Pillow`:

```bash
pip install openpyxl Pillow
```

For pre-evaluated formula values (so previewers that don't recalculate show real numbers in the Standard Path totals), install LibreOffice — the writer will use it headlessly if found:

```bash
sudo apt-get install libreoffice-calc       # Debian/Ubuntu
brew install --cask libreoffice             # macOS
```

If LibreOffice isn't available the xlsx still opens correctly in Excel (which recalculates on open) — only third-party previewers see `#VALUE!` until then.

## data.json shape

```jsonc
{
  "program_code": "MSCSIA",
  "model": { /* program / program_outcomes / ccts / courses (with .competencies) */ },
  "alignments": {
    "course_po":  [{"prog_course_id": "i-...", "po_id":  "i-...", "letters": ["I","R","M","A","X"]}],
    "course_cct": [{"prog_course_id": "i-...", "cct_id": "i-...", "letters": [...]}],
    "comp_po":    [{"pcc_id": "i-...", "course_base_id": "i-...", "po_id":  "i-...", "letters": [...]}],
    "comp_cct":   [{"pcc_id": "i-...", "course_base_id": "i-...", "cct_id": "i-...", "letters": [...]}]
  }
}
```

`letters` is the IRMA-X array as fetched from the Coda junction tables. The merger normalizes them via `letters_to_irma()` (canonical I→R→M→A→X order, dedup'd) before they hit the xlsx.

## Files

| File | Source / role |
|---|---|
| `run.py` | Entry point — argparse, file I/O, calls merger + writer |
| `merger.py` | Pure JSON-driven merger (refactor of `pdow_exporter/merge_and_emit_v16.py`) |
| `xlsx_writer.py` | Snapshot of `pdow_exporter/xlsx_writer.py` — keep in sync if that file changes |
| `template_BSSCOM.xlsx` | Workbook template (snapshot) |
| `owl.png` | Banner logo (snapshot) |
| `data.json` | Bundled MSCSIA v16 sample for smoke-testing the kit |

The three "snapshot" files are duplicates of their counterparts under `pdow_exporter/`. The web UI's bundler is responsible for copying fresh versions in at zip time so the deployed kit is always current.

## Smoke test

```bash
python run.py
# [MERGE] Applied alignments to 11/11 courses
# [MERGE] Applied alignments to 40/40 competencies
# [XLSX] Wrote output/MSCSIA_Program_Map.xlsx in ~8s
```

The resulting xlsx should be byte-identical (modulo timestamps in `docProps/core.xml`) to `output/MSCSIA_Program_Map_v16_live.xlsx` produced by `pdow_exporter/merge_and_emit_v16.py`.
