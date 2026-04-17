"""
PDOW Program Map extractor.

Given a program row ID and a CodaReadClient, fetches all the data needed
to reconstruct the Program Map Excel tab and returns a dict that mirrors
the tab's structure.

Output model shape:
{
  "program": {"id": "i--VmcaKTylm", "name": "M.S. Cybersec & Information Assurance"},
  "program_outcomes": [
      {"id": "i-4MSH5Bo8aN", "name": "Cybersecurity Strategy and Risk Management"},
      ...
  ],
  "ccts": [
      {"id": "i-GOwzoAE6nA", "name": "Risk Management and Compliance"},
      ...
  ],
  "courses": [
      {
        "prog_course_id": "i-5KvaluNhGb",
        "code": "E123",             # from _Courses base
        "name": "Cybersecurity Fundamentals",  # short, from _Courses base
        "display_name": "E123: Cybersecurity Fundamentals",  # from prog_course ref
        "term": "1",
        "std_path_order": 1,
        "cu": 2,
        "scope": "Redesign",
        "ownership": "School of Technology",
        "designation": "Major",
        "certificate": "",
        "po_alignments": {"i-4MSH5Bo8aN": "I, R", ...},   # letters per PO id
        "cct_alignments": {"i-GOwzoAE6nA": "I", ...},     # letters per CCT id
        "competencies": [
            {
              "pcc_id": "i-1uW8QiNqLe",
              "title": "Applies Modern Cybersecurity Principles",
              "statement": "The learner...",
              "level": 2,
              "assessment": "PA",
              "po_alignments": {"i-4MSH5Bo8aN": "X", ...},
              "cct_alignments": {"i-GOwzoAE6nA": "X", ...},
            },
            ...
        ]
      },
      ... sorted by (term_int, std_path_order_or_inf, code)
  ]
}
"""
from coda_client import CodaReadClient
import coda_schema as S


def extract_program_map(doc_id: str, program_id: str, client: CodaReadClient) -> dict:
    """End-to-end extractor. Makes 4-6 table reads (each paginated)."""

    # ----- 1. Program basics -----
    programs = client.read_all_rows(S.doc_table_uri(doc_id, "programs"))
    program_row = next((r for r in programs if r["rowId"] == program_id), None)
    if not program_row:
        raise ValueError(f"Program {program_id} not found in {doc_id}")
    program_name = (
        client.text(program_row["values"].get(S.PROGRAMS_DISPLAY, {}))
        or client.text(program_row["values"].get(S.PROGRAMS_NAME_TEXT, {}))
    )

    # ----- 2. Courses for this program -----
    pc_rows = client.read_all_rows(S.doc_table_uri(doc_id, "prog_courses"))
    mscsia_pc_rows = [
        r for r in pc_rows
        if client.ref_identifier(r["values"].get(S.PC_PROGRAM, {})) == program_id
    ]

    # We need course code/name from _Courses base. Build a lookup.
    courses_rows = client.read_all_rows(S.doc_table_uri(doc_id, "courses"))
    course_base_by_id = {
        r["rowId"]: {
            "code": client.text(r["values"].get(S.COURSE_CODE, {})),
            "name": client.text(r["values"].get(S.COURSE_NAME, {})),
            "cu":   client.number(r["values"].get(S.COURSE_CU, {})),
        }
        for r in courses_rows
    }

    courses = []
    for r in mscsia_pc_rows:
        vals = r["values"]
        course_ref = vals.get(S.PC_COURSE, {})
        course_base_id = client.ref_identifier(course_ref)
        base = course_base_by_id.get(course_base_id, {})
        courses.append({
            "prog_course_id": r["rowId"],
            "course_base_id": course_base_id,
            "code": base.get("code", ""),
            "name": base.get("name", ""),
            "display_name": client.ref_name(course_ref) or "",
            "term": client.text(vals.get(S.PC_TERM, {})),
            "std_path_order": client.number(vals.get(S.PC_STD_PATH, {})),
            "cu": client.number(vals.get(S.PC_CU, {})) or base.get("cu"),
            "scope": client.ref_name(vals.get(S.PC_SCOPE, {})) or "",
            "ownership": client.ref_name(vals.get(S.PC_OWNERSHIP, {})) or "",
            "designation": client.text(vals.get(S.PC_DESIGNATION, {})),
            "certificate": client.text(vals.get(S.PC_CERTIFICATE, {})),
            "scope_notes": client.text(vals.get(S.PC_SCOPE_NOTES, {})),
            "description": client.text(vals.get(S.PC_DESCRIPTION, {})),
            "po_alignments": {},
            "cct_alignments": {},
            "competencies": [],
        })

    # Sort by (term as int, std_path as number, code) — blanks sort last.
    def sort_key(c):
        t = c["term"]
        t_int = int(t) if t.isdigit() else 99
        sp = c["std_path_order"] if c["std_path_order"] is not None else 9999
        return (t_int, sp, c["code"] or "zzz")
    courses.sort(key=sort_key)

    courses_by_pc_id = {c["prog_course_id"]: c for c in courses}

    # ----- 3. Program Outcomes for this program -----
    po_rows = client.read_all_rows(S.doc_table_uri(doc_id, "program_outcomes"))
    program_outcomes = []
    for r in po_rows:
        if client.ref_identifier(r["values"].get(S.PO_PROGRAM, {})) != program_id:
            continue
        slate = r["values"].get(S.PO_NAME_SLATE, {}).get("content", {})
        name = S.extract_po_name_from_slate(slate)
        description = S.extract_po_description_from_slate(slate)
        program_outcomes.append({"id": r["rowId"], "name": name, "description": description})

    # ----- 4. CCTs for this program -----
    cct_rows = client.read_all_rows(S.doc_table_uri(doc_id, "ccts"))
    ccts = []
    for r in cct_rows:
        if client.ref_identifier(r["values"].get(S.CCT_PROGRAM, {})) != program_id:
            continue
        ccts.append({
            "id": r["rowId"],
            "name": client.text(r["values"].get(S.CCT_NAME, {})),
            "description": client.text(r["values"].get(S.CCT_DESCRIPTION, {})),
        })

    # ----- 5. Competencies (PCC) for this program's courses -----
    pcc_rows = client.read_all_rows(S.doc_table_uri(doc_id, "pcc"))
    mscsia_pccs = [
        r for r in pcc_rows
        if client.ref_identifier(r["values"].get(S.PCC_PROGRAM, {})) == program_id
    ]

    pcc_by_id = {}  # rowId -> dict (for later alignment join)
    for r in mscsia_pccs:
        vals = r["values"]
        prog_course_id = client.ref_identifier(vals.get(S.PCC_PROG_COURSE, {}))
        # Fallback: use PCC_COURSE (course base ref) → find matching prog_course
        if not prog_course_id:
            course_base_id = client.ref_identifier(vals.get(S.PCC_COURSE, {}))
            for c in courses:
                if c["course_base_id"] == course_base_id:
                    prog_course_id = c["prog_course_id"]
                    break

        comp = {
            "pcc_id": r["rowId"],
            "prog_course_id": prog_course_id,
            "title": client.text(vals.get(S.PCC_COMP_TITLE, {})),
            "statement": client.text(vals.get(S.PCC_COMP_STMT, {})),
            "level": client.number(vals.get(S.PCC_LEVEL, {})),
            "assessment": client.ref_name(vals.get(S.PCC_ASSESSMENT, {})) or "",
            "po_alignments": {},
            "cct_alignments": {},
        }
        pcc_by_id[r["rowId"]] = comp
        if prog_course_id in courses_by_pc_id:
            courses_by_pc_id[prog_course_id]["competencies"].append(comp)

    # ----- 6. Course × PO alignments -----
    course_x_po = client.read_all_rows(S.doc_table_uri(doc_id, "course_x_po"))
    po_ids = {po["id"] for po in program_outcomes}
    for r in course_x_po:
        vals = r["values"]
        pc_id = client.ref_identifier(vals.get(S.CXA_PROG_COURSE, {}))
        if pc_id not in courses_by_pc_id:
            continue
        po_id = client.ref_identifier(vals.get(S.CXA_TARGET, {}))
        if po_id not in po_ids:
            continue
        letters = _letters_from_aligned(client.array_refs(vals.get(S.CXA_ALIGNED, {})))
        if letters:
            courses_by_pc_id[pc_id]["po_alignments"][po_id] = letters

    # ----- 7. Course × CCT alignments -----
    course_x_cct = client.read_all_rows(S.doc_table_uri(doc_id, "course_x_cct"))
    cct_ids = {c["id"] for c in ccts}
    for r in course_x_cct:
        vals = r["values"]
        pc_id = client.ref_identifier(vals.get(S.CXA_PROG_COURSE, {}))
        if pc_id not in courses_by_pc_id:
            continue
        cct_id = client.ref_identifier(vals.get(S.CXA_TARGET, {}))
        if cct_id not in cct_ids:
            continue
        letters = _letters_from_aligned(client.array_refs(vals.get(S.CXA_ALIGNED, {})))
        if letters:
            courses_by_pc_id[pc_id]["cct_alignments"][cct_id] = letters

    # ----- 8. Competency × PO alignments -----
    comp_x_po = client.read_all_rows(S.doc_table_uri(doc_id, "comp_x_po"))
    for r in comp_x_po:
        vals = r["values"]
        pcc_id = client.ref_identifier(vals.get(S.CMX_PCC, {}))
        if pcc_id not in pcc_by_id:
            continue
        po_id = client.ref_identifier(vals.get(S.CMX_TARGET, {}))
        if po_id not in po_ids:
            continue
        letters = _letters_from_aligned(client.array_refs(vals.get(S.CMX_ALIGNED, {})))
        if letters:
            pcc_by_id[pcc_id]["po_alignments"][po_id] = letters

    # ----- 9. Competency × CCT alignments -----
    comp_x_cct = client.read_all_rows(S.doc_table_uri(doc_id, "comp_x_cct"))
    for r in comp_x_cct:
        vals = r["values"]
        pcc_id = client.ref_identifier(vals.get(S.CMX_PCC, {}))
        if pcc_id not in pcc_by_id:
            continue
        cct_id = client.ref_identifier(vals.get(S.CMX_TARGET, {}))
        if cct_id not in cct_ids:
            continue
        letters = _letters_from_aligned(client.array_refs(vals.get(S.CMX_ALIGNED, {})))
        if letters:
            pcc_by_id[pcc_id]["cct_alignments"][cct_id] = letters

    return {
        "program": {"id": program_id, "name": program_name},
        "program_outcomes": program_outcomes,
        "ccts": ccts,
        "courses": courses,
    }


def _letters_from_aligned(ref_list: list[dict]) -> str:
    """
    Turn [{name: 'I', id: ...}, {name: 'R', id: ...}] into 'I, R'.
    Filters out empty names, preserves order.
    """
    letters = [r.get("name", "").strip() for r in ref_list]
    letters = [l for l in letters if l]
    return ", ".join(letters)
