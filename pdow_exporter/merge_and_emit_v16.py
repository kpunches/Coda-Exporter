#!/usr/bin/env python3
"""v16 alignment merger: takes the 4 junction-table dumps (course_x_po, course_x_cct,
comp_x_po, comp_x_cct) captured via server-side [Program Abbreviat(ed|ion)] filter,
joins into mscsia_model_live.json, writes final xlsx."""
from __future__ import annotations
import json, sys, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[0]))

# Paths are resolved relative to this file's location so the script runs
# from any working directory (repo root, pdow_exporter/, etc.)
HERE = Path(__file__).resolve().parent
DATA = HERE / "data" / "live"
OUT = (HERE.parent / "output")
OUT.mkdir(parents=True, exist_ok=True)

# ---------- Raw alignment data (pasted from MCP tool_results this turn) ---------- #
# We embed only the minimal fields we need: prog_course_id, po/cct id, IRMA letter array.
# This avoids re-fetching or serializing the full JSON.

# course_x_po (55 rows): (prog_course_id, po_id) → letters list
COURSE_PO_RAW: list[tuple[str,str,list[str]]] = [
    # E121
    ("i-ytWQleCouz","i-4MSH5Bo8aN",["A","M"]),
    ("i-ytWQleCouz","i-AtlW5iPFg-",[]),
    ("i-ytWQleCouz","i-s38p6SqZpY",[]),
    ("i-ytWQleCouz","i-BKbajtfrlD",[]),
    ("i-ytWQleCouz","i-l_Maz_Qi3P",["M"]),
    # E123
    ("i-5KvaluNhGb","i-4MSH5Bo8aN",["I"]),
    ("i-5KvaluNhGb","i-AtlW5iPFg-",["I"]),
    ("i-5KvaluNhGb","i-s38p6SqZpY",["I"]),
    ("i-5KvaluNhGb","i-BKbajtfrlD",["I"]),
    ("i-5KvaluNhGb","i-l_Maz_Qi3P",["I"]),
    # D482
    ("i-hBdkIRtsQ8","i-4MSH5Bo8aN",[]),
    ("i-hBdkIRtsQ8","i-AtlW5iPFg-",["I"]),
    ("i-hBdkIRtsQ8","i-s38p6SqZpY",[]),
    ("i-hBdkIRtsQ8","i-BKbajtfrlD",["I"]),
    ("i-hBdkIRtsQ8","i-l_Maz_Qi3P",[]),
    # D483
    ("i-PGzFHOv9vx","i-4MSH5Bo8aN",["R"]),
    ("i-PGzFHOv9vx","i-AtlW5iPFg-",[]),
    ("i-PGzFHOv9vx","i-s38p6SqZpY",[]),
    ("i-PGzFHOv9vx","i-BKbajtfrlD",["A","R"]),
    ("i-PGzFHOv9vx","i-l_Maz_Qi3P",[]),
    # D484
    ("i-hl6pX9EWfq","i-4MSH5Bo8aN",[]),
    ("i-hl6pX9EWfq","i-AtlW5iPFg-",[]),
    ("i-hl6pX9EWfq","i-s38p6SqZpY",[]),
    ("i-hl6pX9EWfq","i-BKbajtfrlD",["M"]),
    ("i-hl6pX9EWfq","i-l_Maz_Qi3P",[]),
    # D485
    ("i-TYNmNslsjU","i-4MSH5Bo8aN",[]),
    ("i-TYNmNslsjU","i-AtlW5iPFg-",["I"]),
    ("i-TYNmNslsjU","i-s38p6SqZpY",["I"]),
    ("i-TYNmNslsjU","i-BKbajtfrlD",[]),
    ("i-TYNmNslsjU","i-l_Maz_Qi3P",[]),
    # D487
    ("i-N-XwVO_Kie","i-4MSH5Bo8aN",[]),
    ("i-N-XwVO_Kie","i-AtlW5iPFg-",[]),
    ("i-N-XwVO_Kie","i-s38p6SqZpY",["A","R"]),
    ("i-N-XwVO_Kie","i-BKbajtfrlD",[]),
    ("i-N-XwVO_Kie","i-l_Maz_Qi3P",[]),
    # D488
    ("i-G1IWjbNavq","i-4MSH5Bo8aN",["R"]),
    ("i-G1IWjbNavq","i-AtlW5iPFg-",["A","R"]),
    ("i-G1IWjbNavq","i-s38p6SqZpY",["R"]),
    ("i-G1IWjbNavq","i-BKbajtfrlD",["R"]),
    ("i-G1IWjbNavq","i-l_Maz_Qi3P",[]),
    # D489
    ("i-1h7bZ1YEi5","i-4MSH5Bo8aN",["M"]),
    ("i-1h7bZ1YEi5","i-AtlW5iPFg-",[]),
    ("i-1h7bZ1YEi5","i-s38p6SqZpY",["M"]),
    ("i-1h7bZ1YEi5","i-BKbajtfrlD",[]),
    ("i-1h7bZ1YEi5","i-l_Maz_Qi3P",["R"]),
    # D490
    ("i-20bcoVkESf","i-4MSH5Bo8aN",["M"]),
    ("i-20bcoVkESf","i-AtlW5iPFg-",["M"]),
    ("i-20bcoVkESf","i-s38p6SqZpY",["M"]),
    ("i-20bcoVkESf","i-BKbajtfrlD",["M"]),
    ("i-20bcoVkESf","i-l_Maz_Qi3P",["M"]),
    # E122
    ("i-3Gk2-onb2E","i-4MSH5Bo8aN",["M"]),
    ("i-3Gk2-onb2E","i-AtlW5iPFg-",[]),
    ("i-3Gk2-onb2E","i-s38p6SqZpY",["M"]),
    ("i-3Gk2-onb2E","i-BKbajtfrlD",["R"]),
    ("i-3Gk2-onb2E","i-l_Maz_Qi3P",["A","M"]),
]

# course_x_cct (44 rows): (prog_course_id, cct_id) → letters list
COURSE_CCT_RAW: list[tuple[str,str,list[str]]] = [
    # E123
    ("i-5KvaluNhGb","i-GOwzoAE6nA",["I"]),
    ("i-5KvaluNhGb","i-VpmGsF6fdL",["I"]),
    ("i-5KvaluNhGb","i-60dxS64cyf",["I"]),
    ("i-5KvaluNhGb","i-bKEkKFt1HO",["I"]),
    # D482
    ("i-hBdkIRtsQ8","i-GOwzoAE6nA",[]),
    ("i-hBdkIRtsQ8","i-VpmGsF6fdL",["R"]),
    ("i-hBdkIRtsQ8","i-60dxS64cyf",["R"]),
    ("i-hBdkIRtsQ8","i-bKEkKFt1HO",[]),
    # D487
    ("i-N-XwVO_Kie","i-GOwzoAE6nA",[]),
    ("i-N-XwVO_Kie","i-VpmGsF6fdL",["R"]),
    ("i-N-XwVO_Kie","i-60dxS64cyf",[]),
    ("i-N-XwVO_Kie","i-bKEkKFt1HO",[]),
    # D483
    ("i-PGzFHOv9vx","i-GOwzoAE6nA",[]),
    ("i-PGzFHOv9vx","i-VpmGsF6fdL",["R"]),
    ("i-PGzFHOv9vx","i-60dxS64cyf",["R"]),
    ("i-PGzFHOv9vx","i-bKEkKFt1HO",["R"]),
    # D485
    ("i-TYNmNslsjU","i-GOwzoAE6nA",["R"]),
    ("i-TYNmNslsjU","i-VpmGsF6fdL",["R"]),
    ("i-TYNmNslsjU","i-60dxS64cyf",["R"]),
    ("i-TYNmNslsjU","i-bKEkKFt1HO",[]),
    # D484
    ("i-hl6pX9EWfq","i-GOwzoAE6nA",[]),
    ("i-hl6pX9EWfq","i-VpmGsF6fdL",["R"]),
    ("i-hl6pX9EWfq","i-60dxS64cyf",["R"]),
    ("i-hl6pX9EWfq","i-bKEkKFt1HO",["R"]),
    # D488
    ("i-G1IWjbNavq","i-GOwzoAE6nA",["R"]),
    ("i-G1IWjbNavq","i-VpmGsF6fdL",["R"]),
    ("i-G1IWjbNavq","i-60dxS64cyf",["R"]),
    ("i-G1IWjbNavq","i-bKEkKFt1HO",["R"]),
    # E121
    ("i-ytWQleCouz","i-GOwzoAE6nA",["R"]),
    ("i-ytWQleCouz","i-VpmGsF6fdL",["R"]),
    ("i-ytWQleCouz","i-60dxS64cyf",["R"]),
    ("i-ytWQleCouz","i-bKEkKFt1HO",["R"]),
    # D489
    ("i-1h7bZ1YEi5","i-GOwzoAE6nA",["R"]),
    ("i-1h7bZ1YEi5","i-VpmGsF6fdL",["R"]),
    ("i-1h7bZ1YEi5","i-60dxS64cyf",["R"]),
    ("i-1h7bZ1YEi5","i-bKEkKFt1HO",["R"]),
    # D490
    ("i-20bcoVkESf","i-GOwzoAE6nA",["M"]),
    ("i-20bcoVkESf","i-VpmGsF6fdL",["M"]),
    ("i-20bcoVkESf","i-60dxS64cyf",["M"]),
    ("i-20bcoVkESf","i-bKEkKFt1HO",["M"]),
    # E122
    ("i-3Gk2-onb2E","i-GOwzoAE6nA",[]),
    ("i-3Gk2-onb2E","i-VpmGsF6fdL",["R"]),
    ("i-3Gk2-onb2E","i-60dxS64cyf",["R"]),
    ("i-3Gk2-onb2E","i-bKEkKFt1HO",["R"]),
]

# comp_x_po (200 rows): (pcc_id, po_id, course_id, competency_id, letters[])
# We key competency alignments by PCC id because comp_x_po references PCC not just competency
COMP_PO_RAW: list[tuple[str,str,str,str,list[str]]] = []

# comp_x_cct (160 rows): same shape
COMP_CCT_RAW: list[tuple[str,str,str,str,list[str]]] = []

def letters_to_irma(letters: list[str]) -> str:
    """Join IRMA letters in canonical order."""
    order = ["I","R","M","A","X"]
    seen = set(letters)
    return "".join(L for L in order if L in seen)

def main():
    t0 = time.time()
    model_path = DATA / "mscsia_model_live.json"
    model = json.loads(model_path.read_text())

    # Build lookup maps
    po_order = [po["id"] for po in model["program_outcomes"]]
    cct_order = [c["id"] for c in model["ccts"]]

    # Populate course.po_alignments and course.cct_alignments
    course_po_map: dict[str, dict[str,str]] = {}
    for pc_id, po_id, letters in COURSE_PO_RAW:
        course_po_map.setdefault(pc_id, {})[po_id] = letters_to_irma(letters)

    course_cct_map: dict[str, dict[str,str]] = {}
    for pc_id, cct_id, letters in COURSE_CCT_RAW:
        course_cct_map.setdefault(pc_id, {})[cct_id] = letters_to_irma(letters)

    matched = 0
    for course in model.get("courses", []):
        pc_id = course.get("prog_course_id") or course.get("id")
        pa = course_po_map.get(pc_id, {})
        ca = course_cct_map.get(pc_id, {})
        course["po_alignments"] = {po: pa.get(po,"") for po in po_order}
        course["cct_alignments"] = {c: ca.get(c,"") for c in cct_order}
        if pa or ca:
            matched += 1
    print(f"[MERGE] Applied alignments to {matched}/{len(model.get('courses',[]))} courses")

    # Now competency alignments, keyed by (course_base_id, pcc_id)
    sys.path.insert(0, str(DATA / "alignments_v2"))
    from comp_alignments_data import COMP_PO, COMP_CCT

    comp_po_map: dict[tuple[str,str], dict[str,str]] = {}
    for pcc_id, course_base, po_id, letters in COMP_PO:
        comp_po_map.setdefault((course_base, pcc_id), {})[po_id] = letters_to_irma(letters)

    comp_cct_map: dict[tuple[str,str], dict[str,str]] = {}
    for pcc_id, course_base, cct_id, letters in COMP_CCT:
        comp_cct_map.setdefault((course_base, pcc_id), {})[cct_id] = letters_to_irma(letters)

    comp_matched = 0
    comp_total = 0
    for course in model.get("courses", []):
        cb = course.get("course_base_id")
        for comp in course.get("competencies", []):
            comp_total += 1
            pcc = comp.get("pcc_id")
            pa = comp_po_map.get((cb, pcc), {})
            ca = comp_cct_map.get((cb, pcc), {})
            comp["po_alignments"] = {po: pa.get(po,"") for po in po_order}
            comp["cct_alignments"] = {c: ca.get(c,"") for c in cct_order}
            if pa or ca:
                comp_matched += 1
    print(f"[MERGE] Applied alignments to {comp_matched}/{comp_total} competencies")

    # Save updated model
    out_model = DATA / "mscsia_model_v16.json"
    out_model.write_text(json.dumps(model, indent=2))
    print(f"[SAVE] Wrote model → {out_model}")

    # Write XLSX using existing writer (write_program_map takes explicit paths)
    from xlsx_writer import write_program_map
    template = HERE / "template_BSSCOM.xlsx"
    owl = HERE / "owl.png"
    out_xlsx = OUT / "MSCSIA_Program_Map_v16_live.xlsx"
    write_program_map(
        model=model,
        template_path=template,
        output_path=out_xlsx,
        owl_image_path=owl,
        tab_name="Program Map",
        program_code="MSCSIA",
    )
    print(f"[XLSX] Wrote {out_xlsx} in {time.time()-t0:.1f}s")
    return out_xlsx

if __name__ == "__main__":
    main()
