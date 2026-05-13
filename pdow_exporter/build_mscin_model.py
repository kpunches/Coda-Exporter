"""
Build mscin_model.json from the raw fetched data under data/live/mscin/.

Mirrors the v16 MSCSIA model shape, with one addition:
  course["competencies"][i]["skills"] = [
      {"id": "...", "category": "...", "title": "...",
       "statement": "...", "url": "..."}, ...
  ]

CCT alignments are intentionally left blank for the MSCIN run per Keith's
instruction (CCT IRMA mapping has not been populated yet).
"""
from __future__ import annotations

import json
from pathlib import Path

LIVE = Path(__file__).parent / "data" / "live" / "mscin"
OUT = Path(__file__).parent / "data" / "live" / "mscin_model.json"

PROGRAM_ID = "i-ws6bn0g4BY"
PROGRAM_NAME = "M.S. Curriculum and Instruction"


def _load(name: str):
    with (LIVE / name).open() as f:
        return json.load(f)


def _text(cell: dict) -> str:
    content = cell.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, (int, float)):
        return str(content)
    if isinstance(content, dict):
        if content.get("type") == "num":
            return str(content.get("value", ""))
    return ""


def _num(cell: dict):
    content = cell.get("content")
    if isinstance(content, (int, float)):
        return content
    if isinstance(content, dict) and content.get("type") == "num":
        return content.get("value")
    return None


def _ref_id(cell: dict):
    content = cell.get("content")
    if isinstance(content, dict) and content.get("type") == "ref":
        return content.get("identifier")
    return None


def _ref_name(cell: dict):
    content = cell.get("content")
    if isinstance(content, dict) and content.get("type") == "ref":
        return content.get("name")
    return None


def _short_school(school_name: str) -> str:
    return {
        "School of Business": "WSB",
        "School of Technology": "WST",
        "School of Health Leavitt": "WSH",
        "School of Education": "WSE",
        "General Education": "Gen Ed",
    }.get(school_name, school_name)


def build():
    # ---- Programs / POs / CCTs / courses base ----
    pos = _load("program_outcomes.json")
    ccts = _load("ccts.json")
    courses_base = {c["id"]: c for c in _load("courses_base.json")}

    # ---- Courses (prog_courses) ----
    pc_raw = _load("prog_courses.json")["rows"]
    courses = []
    for r in pc_raw:
        vals = r["values"]
        course_base_id = _ref_id(vals.get("c-aGpuFk4ifn", {}))
        base = courses_base.get(course_base_id, {})
        term_val = _num(vals.get("c-j0Y6lwi_yU", {}))
        term = str(int(term_val)) if isinstance(term_val, (int, float)) else (term_val or "")
        sp_val = _num(vals.get("c-tJdYNtbcRy", {}))
        courses.append({
            "prog_course_id": r["rowId"],
            "course_base_id": course_base_id,
            "code": base.get("code", ""),
            "name": base.get("name", ""),
            "display_name": _ref_name(vals.get("c-aGpuFk4ifn", {})) or "",
            "term": term,
            "std_path_order": sp_val,
            "cu": _num(vals.get("c-DTsXuT3z72", {})) or base.get("cu"),
            "scope": _ref_name(vals.get("c-UlpI6zHbtU", {})) or "",
            "ownership": _short_school(_ref_name(vals.get("c-SSH1vm1eG5", {})) or ""),
            "designation": _text(vals.get("c-YZLH4IcD51", {})),
            "certificate": _text(vals.get("c-sj2stjNrbN", {})),
            "scope_notes": _text(vals.get("c-nHeKFgF3C5", {})),
            "description": _text(vals.get("c-PoBDbaSL6S", {})),
            "po_alignments": {},
            "cct_alignments": {},
            "competencies": [],
        })

    def _sort_key(c):
        t = c["term"]
        t_int = int(t) if str(t).isdigit() else 99
        sp = c["std_path_order"] if c["std_path_order"] is not None else 9999
        return (t_int, sp, c["code"] or "zzz")
    courses.sort(key=_sort_key)

    courses_by_pc = {c["prog_course_id"]: c for c in courses}

    # ---- Competencies (PCC) ----
    pcc_data = _load("pcc.json")["rows"]
    comp_by_id = {}
    for r in pcc_data:
        comp = {
            "pcc_id": r["rowId"],
            "prog_course_id": r["prog_course"],
            "title": r["title"],
            "statement": r["statement"],
            "level": r["level"],
            "assessment": r["assessment"],
            "po_alignments": {},
            "cct_alignments": {},
            "skills": [],
            # comp_order is the position within the course (filled by iteration below)
            "comp_order": 0,
        }
        comp_by_id[r["rowId"]] = comp
        if comp["prog_course_id"] in courses_by_pc:
            courses_by_pc[comp["prog_course_id"]]["competencies"].append(comp)

    # Stamp comp_order within each course (1-based, preserving extractor order)
    for c in courses:
        for i, comp in enumerate(c["competencies"], start=1):
            comp["comp_order"] = i

    # ---- Course × PO alignments ----
    for pc_id, po_id, letters in _load("course_x_po.json"):
        if pc_id in courses_by_pc:
            courses_by_pc[pc_id]["po_alignments"][po_id] = letters

    # ---- Competency × PO alignments ----
    for pcc_id, po_id, letters in _load("comp_x_po.json"):
        if pcc_id in comp_by_id:
            comp_by_id[pcc_id]["po_alignments"][po_id] = letters

    # CCT alignments intentionally blank (per user instruction).

    # ---- Skills ----
    skills_data = _load("comp_x_skill.json")
    for s in skills_data:
        pcc_id = s["pcc_id"]
        if pcc_id not in comp_by_id:
            continue
        comp_by_id[pcc_id]["skills"].append({
            "id": s["skill_id"],
            "category": s["category"] or "",
            "title": s["title"] or "",
            "statement": s["statement"] or "",
            "url": s["url"] or "",
        })

    # Sort skills within each comp by category, then title
    for comp in comp_by_id.values():
        comp["skills"].sort(key=lambda sk: (sk["category"].lower(), sk["title"].lower()))

    model = {
        "program": {"id": PROGRAM_ID, "name": PROGRAM_NAME},
        "program_outcomes": pos,
        "ccts": ccts,
        "courses": courses,
    }

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w") as f:
        json.dump(model, f, indent=2)

    n_skills = sum(len(c["skills"]) for c in comp_by_id.values())
    print(f"Wrote {OUT}")
    print(f"  courses: {len(courses)}")
    print(f"  competencies: {len(comp_by_id)}")
    print(f"  comp×skill mappings: {n_skills}")
    print(f"  unique skills: {len({s['skill_id'] for s in skills_data})}")


if __name__ == "__main__":
    build()
