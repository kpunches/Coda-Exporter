# Coda-Exporter: PDOW Program Map Builder

Excel exporter that reads program data from Coda doc `4YIajnJqvo` ("Design & Development V3") and produces a fully formatted Program Map workbook (Standard Path, Instructions, Program Map, Course Alignment tabs) with real IRMA-X alignment data.

**Latest shipped:** MSCSIA v16, built 2026-04-17.

---

## If you are Claude Code reading this: start here

1. Read `pdow_exporter/RESUME.md` — it has the full project state, the filter breakthrough notes, and the step-by-step runbook for exporting any program.
2. The Coda MCP connector should be available as tools named like `Coda MCP:table_rows_read`. If it isn't, tell Keith — the whole workflow depends on it.
3. When Keith asks for a new program export (BSN, MSN, etc.), follow the "How to run another program" section in `RESUME.md`.
4. Before writing any Coda `filterFormula`, read about the four silent-failure traps in `RESUME.md` § "The filter breakthrough". The short version: plain-text abbreviation columns are the winning filter target — use `[Program Abbreviated] = "<ABBR>"` for course-level junctions and `[Program Abbreviation] = "<ABBR>"` for comp-level.

---

## Project layout

```
pdow_exporter/
  RESUME.md                   ← handoff doc; read this first
  coda_schema.py              ← immutable table / column / row IDs for doc 4YIajnJqvo
  coda_client.py              ← read client with pagination
  extractor.py                ← extract_program_map(doc_id, program_id, client)
  xlsx_writer.py              ← write_program_map(...); tabs, styling, formula recalc
  merge_and_emit_v16.py       ← the v16 builder; worked example for MSCSIA
  template_BSSCOM.xlsx        ← template the writer copies from
  owl.png                     ← banner logo
  data/
    live/
      mscsia_model_live.json         ← base live model (no alignments)
      mscsia_model_v16.json          ← v16 model with all alignments populated
      alignments_v2/
        _fetch_log.md                ← fetch provenance
        comp_alignments_data.py      ← competency-level alignment tuples (COMP_PO, COMP_CCT)
```

---

## Environment

Python 3.10+ with:

```
pip install openpyxl
```

For formula recalculation on the final xlsx, the writer uses headless LibreOffice:

```
sudo apt-get install libreoffice-calc     # Debian/Ubuntu
brew install --cask libreoffice           # macOS
```

If LibreOffice isn't available, the xlsx still opens correctly — Excel recalculates on open — but the Standard Path term totals will show `#VALUE!` until then. LibreOffice headless is the cleanest fix.

---

## Running MSCSIA (repro the v16 build)

```bash
cd pdow_exporter
python merge_and_emit_v16.py
# → writes ../output/MSCSIA_Program_Map_v16_live.xlsx (or wherever OUT is set)
```

The merger uses the already-fetched alignment data baked into `comp_alignments_data.py` and the course-level tuples in `merge_and_emit_v16.py` itself, so no Coda calls are needed to reproduce the MSCSIA build.

## Running a new program (BSN, MSN, etc.)

See `pdow_exporter/RESUME.md` § "How to run another program" for the full step-by-step. Short version:

1. Get the program's abbreviation and row ID from the programs table (`grid-_cLrawcUzd`)
2. Build the base model via `extractor.py`
3. Fetch the four junction tables with server-side filter on `[Program Abbreviated]` / `[Program Abbreviation]`
4. Merge alignments into the model using the patterns in `merge_and_emit_v16.py`
5. Call `xlsx_writer.write_program_map(...)`

---

## Known gotchas

- **Coda filter silent failures.** If `filterFormula` references an unknown column or uses a wrong idiom for the column type, `table_rows_read` returns rows anyway and puts the error in `result.filterFormulaError`. Always check that field.
- **`.ID()` is not a CFL function.** Use `RowId()`.
- **Table back-references aren't filterable.** On some junction tables a pipe-named field (e.g. `_Progs | _Courses`) is a table back-reference, not a scalar. Pick a different filter column.
- **Same column ID, different display name per table.** CFL resolves by name. Check the schema per table with `document_read`.

All four traps are documented in detail in `RESUME.md` and in the `coda-mcp-patterns` skill (if you're running this under a Claude environment that loads that skill).

---

## What was built in the original Claude.ai sessions

Transcripts for the multi-session build live on the Claude.ai filesystem at `/mnt/transcripts/`. The key ones (newest first):

- `2026-04-17-02-21-54-pdow-mscsia-v16-filter-breakthrough.txt` — the definitive filter fix + v16 shipping
- `2026-04-17-02-18-42-pdow-mscsia-v16-breakthrough.txt` — mid-execution, course_x_po fetched
- `2026-04-17-01-53-57-pdow-exporter-live-run.txt` — RegexMatch + ToText breakthrough
- `2026-04-17-01-30-41-pdow-exporter-build.txt` — writer build-out
- `2026-04-16-22-44-27-pdow-excel-exporter-build.txt` — original kick-off

Those aren't in this repo; they're for reference if you need to trace a decision's history.
