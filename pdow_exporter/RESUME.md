# PDOW Exporter — RESUME

**Last updated:** 2026-04-17 (v16 delivery + repo packaging + Windows fd fix)
**User:** Keith (WGU instructional design). Working in Coda doc `4YIajnJqvo` ("Design & Development V3").

---

## If you are a fresh Claude reading this: start here

1. Read this file top to bottom — it's short.
2. Figure out which environment you're in:
   - **Claude Code (local repo):** project root is wherever you were started. All code is here. You have the Coda MCP connector for live data fetches.
   - **Claude.ai chat session:** this repo was zipped and dropped onto the sandbox. Run `ls` to locate it.
3. If Keith asks you to do anything involving Coda junction table filters, read the "Filter breakthrough" section below before writing any `filterFormula`.
4. If Keith wants a new program export (BSN, MSN, MSCSIA again, etc.), follow the "How to run another program" section.

---

## Current state: MSCSIA v16 shipped ✓

The first fully populated Program Map export — every course row and every competency row has real IRMA-X data in its PO/CCT alignment cells (no blanks like v15).

**Reproducible from this repo:** `python pdow_exporter/merge_and_emit_v16.py` writes `output/MSCSIA_Program_Map_v16_live.xlsx` using the already-fetched alignment data baked into `comp_alignments_data.py` and the course-level tuples in the merger itself. No Coda calls needed for repro.

---

## The filter breakthrough (why it matters for any future program)

Four traps, all hard-learned:

1. **`filterFormulaError` is a silent failure.** `table_rows_read` returns rows even when the filter didn't parse. Always check `result.filterFormulaError` — if populated, the filter was ignored and the data is unfiltered.

2. **`.ID()` is not a CFL function.** Use **`RowId()`**. The Relational sub-guide (`tool_guide` with topic `formula`, sub-topic `Relational`) confirms it.

3. **Table back-references aren't filterable.** On some junction tables, `_Progs | _Courses` in the row drawer is a *table back-reference*, not a scalar column. `thisRow.[TargetTable]` yields "unknown value". Pick a different column via `document_read`.

4. **Same column ID, different name per table.** `c-Dn7elyFRLj` is named `_progs | _courses` on one junction and `Program to Course Pairing` on another. CFL resolves by name, not ID. Always fetch the schema for the specific table.

**The winning pattern** for any program in this doc:

```
filterFormula: [Program Abbreviated] = "<PROG_ABBR>"   # course-level junctions
filterFormula: [Program Abbreviation] = "<PROG_ABBR>"  # comp-level junctions
```

All four junction tables have a precomputed plain-text abbreviation column that's optimal for server-side filtering.

---

## Project files

| File | Purpose |
|---|---|
| `pdow_exporter/coda_schema.py` | Immutable table/column/row IDs for doc 4YIajnJqvo |
| `pdow_exporter/coda_client.py` | Read client with pagination |
| `pdow_exporter/extractor.py` | `extract_program_map(doc_id, program_id, client)` |
| `pdow_exporter/xlsx_writer.py` | `write_program_map(...)` — the main writer. Tabs, styling, formula recalc, VML cleanup |
| `pdow_exporter/merge_and_emit_v16.py` | The v16 builder — merges alignments into the live model and emits xlsx |
| `pdow_exporter/template_BSSCOM.xlsx`, `owl.png` | Template + logo used by writer |
| `pdow_exporter/data/live/mscsia_model_live.json` | Base live model (no alignments) |
| `pdow_exporter/data/live/mscsia_model_v16.json` | v16 model with all alignments populated |
| `pdow_exporter/data/live/alignments_v2/comp_alignments_data.py` | Extracted competency-level alignment tuples (COMP_PO, COMP_CCT) |
| `pdow_exporter/data/live/alignments_v2/_fetch_log.md` | Fetch provenance |

---

## How to run another program (e.g., BSN, MSN, MSCIA)

The hard part is done — filter idioms are known, writer is stable, template is locked. A new-program run is mechanical:

1. **Get the program's abbreviation and row ID** from `grid-_cLrawcUzd` (programs table). Look at the display columns to find how the abbreviation is exposed.

2. **Build the base model** using `extractor.py` with `program_id=<new program row ID>`. This pulls program, POs, CCTs, courses, competencies.

3. **Fetch the four junction tables** with server-side filter:

   | Junction | Grid ID | Abbreviation column | Useful `filterColumnNames` |
   |---|---|---|---|
   | course × PO | `grid-tUkDsBwjVg` | `Program Abbreviated` | `_progs \| _courses`, `Program Learning Outcome`, `Course IRMA Map` |
   | course × CCT | `grid-bdbEJT9Kuq` | `Program Abbreviated` | `Program to Course Pairing`, `Cross Cutting Theme`, `Aligned` |
   | comp × PO | `grid-NkK5JDeDhF` | `Program Abbreviation` | `Program Course Competency`, `Program Outcome`, `Course Comp IRMA Map to PO`, `Course`, `Competency` |
   | comp × CCT | `grid-6xLuZLsQ_t` | `Program Abbreviation` | `Program Course Competency`, `Cross-Cutting Theme`, `Aligned`, `Course`, `Competency` |

4. **Merge into the model** using the pattern in `merge_and_emit_v16.py`:
   - `COURSE_PO_RAW` / `COURSE_CCT_RAW`: list of `(prog_course_id, po_or_cct_id, letters[])`
   - `COMP_PO` / `COMP_CCT`: list of `(pcc_id, course_base_id, po_or_cct_id, letters[])` in `comp_alignments_data.py`
   - Merger applies via `course["po_alignments"][po_id] = letters_to_irma(letters)` and same for competencies.

5. **Call `write_program_map(model, template_path, output_path, owl_image_path, program_code="<ABBR>")`** from `xlsx_writer.py`. It handles tab rebuild, styling, formula recalc (via headless LibreOffice if available), and VML cleanup.

6. **Ship the xlsx** — it'll be in `output/`.

**Don't hand-transcribe tuples.** For MSCSIA we hand-copied 59 course + 400 comp rows into Python lists because the alternative — saving raw JSON from tool responses — wasn't set up. A better pattern: write a `fetch_alignments.py` that saves each `table_rows_read` response to `data/live/alignments_v2/<program>/course_x_po.json` etc., then has the merger parse those files. That's the first refactor candidate.

---

## Environment

Python 3.10+ with `openpyxl` and `Pillow` (Pillow is needed for the owl image embed). For formula recalculation on the final xlsx, the writer uses headless LibreOffice (`libreoffice --headless --calc`). If it's not available, the xlsx still opens — Excel recalculates on open — but Standard Path term totals show `#VALUE!` until then.

**Windows note:** `_strip_stale_vml_comments` in `xlsx_writer.py` previously leaked the temp fd from `tempfile.mkstemp`, which caused `WinError 32` on `shutil.move` at the end of the write. Fixed 2026-04-17 (commit `094406e`) by closing the fd explicitly.

---

## Pending / nice-to-haves

- **Fetch-to-disk helper.** Per the note above — the first real candidate for a refactor. Would also enable a `fetch_program.py <prog_abbr>` CLI that does steps 3-6 end to end.
- **Skill hardening.** A `pdow-exporter` skill scaffold existed in the Claude.ai sandbox but hasn't been promoted. If we want to run this via Claude.ai skills in the future, it should reference `coda-mcp-patterns` for filter idioms rather than re-stating them.
- **Expand the junction schema docs.** Only four junction tables are documented. If Keith builds other junction-style exports (e.g., competency × skill via `_PCC | _RSDs` at `grid-bYtzEaQQ0t`), those should get the same treatment.
- **comp_x_cct is sparse.** All 160 fetched rows for MSCSIA had empty `Aligned` arrays. Not a bug on our side — just how the data is in the doc. Worth confirming with Keith whether that's intentional or pending input.

---

## User context to carry forward

- Keith is actively importing nursing informatics SSDs (completed D951 and D952). The Coda doc is approaching a size threshold that may need architectural attention.
- Coda commenting limitations: cell-level only at row level, no field-level comments in detail views, no Comments column type, etc. "MCP now works with Coda commenting" per the latest update.

---

## Related artifacts (Claude.ai sandbox only — not in this repo)

- Transcripts at `/mnt/transcripts/` on the Claude.ai sandbox, cataloged in `journal.txt`
- Related import skills at `/mnt/skills/user/`: `pdow-coda-mapping`, `ccw-coda-import`, `ssd-coda-import`, `pdvs-coda-import`, `lr-tools-coda-import`, `es-scope-import`, `upload-coordinator`, `extract-qa`, `coda-preflight`, `coda-mcp-patterns`
- The `coda-mcp-patterns` skill got the full filter taxonomy baked in during the v16 session — section "★ CANONICAL: Server-Side Filtering"

None of those are in this repo. They live on the Claude.ai filesystem and are referenced here for cross-environment context.

---

## How to update this file

When you finish a session that materially changed the exporter or the filter knowledge, update this doc:
- Bump "Last updated"
- Add a line under "Pending / nice-to-haves" if you left anything unfinished
- If you shipped a new version, update "Current state" with the new output path

Keep it short. This is a handoff doc, not a changelog.
