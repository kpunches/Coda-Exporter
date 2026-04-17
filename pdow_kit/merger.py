"""
JSON-driven merger for Program Map exports.

Takes a `data` dict with the shape produced by the UI bundler:

    {
      "program_code": "MSCSIA",
      "model": { ...base model: program / program_outcomes / ccts / courses... },
      "alignments": {
        "course_po":  [{"prog_course_id": ..., "po_id":  ..., "letters": [...]}, ...],
        "course_cct": [{"prog_course_id": ..., "cct_id": ..., "letters": [...]}, ...],
        "comp_po":    [{"pcc_id": ..., "course_base_id": ..., "po_id":  ..., "letters": [...]}, ...],
        "comp_cct":   [{"pcc_id": ..., "course_base_id": ..., "cct_id": ..., "letters": [...]}, ...]
      }
    }

Returns the merged model dict (same shape as the base, with `po_alignments` /
`cct_alignments` populated on each course and competency).

Pure logic — no I/O, no Coda calls. Ported from
pdow_exporter/merge_and_emit_v16.py with the hardcoded MSCSIA tuples lifted
out into the input.
"""
from __future__ import annotations


_IRMA_ORDER = ["I", "R", "M", "A", "X"]


def letters_to_irma(letters):
    """Join IRMA letters in canonical order, deduplicated."""
    seen = set(letters or [])
    return "".join(L for L in _IRMA_ORDER if L in seen)


def merge(data):
    """Apply alignments from `data["alignments"]` onto `data["model"]`.

    Mutates and returns the model dict. Prints a one-line summary of how many
    courses and competencies got at least one alignment applied — useful for
    detecting silent miskeys (e.g. wrong prog_course_id format) without
    needing to open the resulting xlsx.
    """
    model = data["model"]
    alignments = data.get("alignments", {})

    po_order  = [po["id"] for po in model.get("program_outcomes", [])]
    cct_order = [c["id"]  for c in model.get("ccts", [])]

    # ── Course-level lookup maps ────────────────────────────────────────────
    course_po_map  = {}
    for row in alignments.get("course_po", []):
        course_po_map.setdefault(row["prog_course_id"], {})[row["po_id"]] = letters_to_irma(row.get("letters"))

    course_cct_map = {}
    for row in alignments.get("course_cct", []):
        course_cct_map.setdefault(row["prog_course_id"], {})[row["cct_id"]] = letters_to_irma(row.get("letters"))

    course_matched = 0
    for course in model.get("courses", []):
        pc_id = course.get("prog_course_id") or course.get("id")
        pa = course_po_map.get(pc_id, {})
        ca = course_cct_map.get(pc_id, {})
        course["po_alignments"]  = {po: pa.get(po,  "") for po in po_order}
        course["cct_alignments"] = {c:  ca.get(c,  "") for c in cct_order}
        if pa or ca:
            course_matched += 1
    print(f"[MERGE] Applied alignments to {course_matched}/{len(model.get('courses', []))} courses")

    # ── Competency-level lookup maps, keyed by (course_base_id, pcc_id) ─────
    comp_po_map  = {}
    for row in alignments.get("comp_po", []):
        key = (row["course_base_id"], row["pcc_id"])
        comp_po_map.setdefault(key, {})[row["po_id"]] = letters_to_irma(row.get("letters"))

    comp_cct_map = {}
    for row in alignments.get("comp_cct", []):
        key = (row["course_base_id"], row["pcc_id"])
        comp_cct_map.setdefault(key, {})[row["cct_id"]] = letters_to_irma(row.get("letters"))

    comp_matched = 0
    comp_total   = 0
    for course in model.get("courses", []):
        cb = course.get("course_base_id")
        for comp in course.get("competencies", []):
            comp_total += 1
            pcc = comp.get("pcc_id")
            pa = comp_po_map.get((cb, pcc), {})
            ca = comp_cct_map.get((cb, pcc), {})
            comp["po_alignments"]  = {po: pa.get(po,  "") for po in po_order}
            comp["cct_alignments"] = {c:  ca.get(c,  "") for c in cct_order}
            if pa or ca:
                comp_matched += 1
    print(f"[MERGE] Applied alignments to {comp_matched}/{comp_total} competencies")

    return model
