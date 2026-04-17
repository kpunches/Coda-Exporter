"""
Coda schema constants for the Design & Development V3 doc.

All IDs are IMMUTABLE — column names can change in the UI, but these IDs
survive. Only updates needed if a column/table is deleted and recreated.

DOC_ID is parameterized via the exporter's runtime args (NOT hardcoded here),
so the same code works against copies of the doc with different doc IDs.
"""

# ----- Table grid IDs -----
TABLES = {
    "programs":         "grid-_cLrawcUzd",   # _Programs
    "courses":          "grid-8i2Q6-eoTP",   # _Courses (base)
    "prog_courses":     "grid-fEebuvIQBl",   # _Progs | _Courses (junction)
    "competencies":     "grid-dXlcRqqYc3",   # _Comps (base)
    "pcc":              "grid-TmGA_WNe3_",   # _Progs_Courses | _Comps
    "program_outcomes": "grid-aFulSXgnJe",   # _Progs Outcomes
    "ccts":             "grid-8Azxi66Mj-",   # _CCTs
    "schools":          "grid-hTPaOOsuXJ",   # Schools/Ownership catalog
    "assessments":      "grid-9UYlau7UHJ",   # OA/PA/etc.
    "irma_values":      "grid-43xRf-GFUo",   # I/R/M/A/X lookup
    # Alignment junction tables (4)
    "course_x_po":      "grid-tUkDsBwjVg",   # _Progs_Courses | _Progs Outcomes
    "course_x_cct":     "grid-bdbEJT9Kuq",   # _PC | _CCTs
    "comp_x_po":        "grid-NkK5JDeDhF",   # _Progs_Courses_Comps | _Progs Outcomes
    "comp_x_cct":       "grid-6xLuZLsQ_t",   # _PCC | _CCTs
}

# ----- _Progs | _Courses columns -----
PC_PROGRAM       = "c-UaE_k1Ivfh"  # ref → programs
PC_COURSE        = "c-aGpuFk4ifn"  # ref → courses (base)
PC_STD_PATH      = "c-tJdYNtbcRy"  # number
PC_TERM          = "c-j0Y6lwi_yU"  # text
PC_CU            = "c-DTsXuT3z72"  # number
PC_SCOPE         = "c-UlpI6zHbtU"  # ref → select list (New/Existing/Redesign)
PC_OWNERSHIP     = "c-SSH1vm1eG5"  # ref → schools
PC_DESIGNATION   = "c-YZLH4IcD51"  # text (Core / Major / GE-... )
PC_CERTIFICATE   = "c-sj2stjNrbN"  # text (usually blank)
PC_SCOPE_NOTES   = "c-nHeKFgF3C5"  # slate — course-level scope notes (New/Redesign)
PC_DESCRIPTION   = "c-PoBDbaSL6S"  # plain text — course description (Existing)

# ----- _Courses (base) columns -----
COURSE_CODE      = "c-sgyJdn2bVc"  # "E123", "Core 1"
COURSE_NAME      = "c-nXrSiX6Q7R"  # short name
COURSE_CU        = "c-8oRBhcidnK"  # number

# ----- _Programs columns -----
PROGRAMS_NAME_TEXT = "c-JbFOWBC9i0"  # plain-text program name (e.g. "Cybersec & Information Assurance")
PROGRAMS_DISPLAY   = "c-L3c8jGheXt"  # "B.S. ..." / "M.S. ..." display form

# ----- PCC (_Progs_Courses | _Comps) columns -----
PCC_PROGRAM      = "c-a4-JBknQCD"  # ref → programs
PCC_COURSE       = "c-UfDWlbwQXC"  # ref → courses (base)
PCC_COMPETENCY   = "c-D-u5Q1gcJp"  # ref → competencies (base)
PCC_PROG_COURSE  = "c-WLMD7Ir-os"  # ref → prog_courses (junction)
PCC_COMP_TITLE   = "c-XRPTQDN2ZN"  # text
PCC_COMP_STMT    = "c-ZQ354aYNfj"  # text
PCC_LEVEL        = "c-FZuaohTac9"  # number
PCC_ASSESSMENT   = "c-GunVQKG_Ir"  # ref → assessments (PA/OA/etc)
PCC_PROG_CODE    = "c-jdZs26VuaI"  # text program code (MSCSIA, BSSCOM, etc.)

# ----- _Progs Outcomes columns -----
PO_NAME_SLATE    = "c-h3bw7zZXzf"  # slate: "Program Outcome: <name>\n<description>"
PO_PROGRAM       = "c-joP6IMYsmR"  # ref → programs

# ----- _CCTs columns -----
CCT_NAME         = "c-eJu3miLkC_"  # text
CCT_DESCRIPTION  = "c-n4q71ORRVT"  # text
CCT_PROGRAM      = "c-joP6IMYsmR"  # ref → programs (same col ID by coincidence)

# ----- Alignment junction columns (Course × PO and Course × CCT) -----
CXA_PROG_COURSE  = "c-Dn7elyFRLj"  # ref → prog_courses
CXA_TARGET       = "c-4CuABuCpCx"  # ref → PO or CCT
CXA_ALIGNED      = "c-bBVinldWlI"  # array of IRMA refs

# ----- Alignment junction columns (Competency × PO and Competency × CCT) -----
CMX_PCC          = "c-pvQqegHRGa"  # ref → PCC
CMX_TARGET       = "c-K0OcZXEwDe"  # ref → PO or CCT
CMX_ALIGNED      = "c-BKtOyobxvu"  # array of IRMA refs


# ----- Expected IRMA/X identifier -> letter mapping (for sanity check / fallback) -----
# In practice, the ref's `.name` field already contains "I"/"R"/"M"/"A"/"X",
# so we extract that directly instead of looking up by ID.
IRMA_NAME_BY_ID = {
    "i-8_5gzmIwd9": "I",
    "i-XkktOH0eLQ": "R",
    "i-R6OoFc4dFh": "M",
    "i-q82R6tJ9rS": "A",
    "i-3z6tPUZmNs": "X",
}


def doc_table_uri(doc_id: str, table_key: str) -> str:
    """Build a coda:// table URI for a given doc and logical table name."""
    grid = TABLES[table_key]
    return f"coda://docs/{doc_id}/tables/{grid}"


def extract_po_name_from_slate(slate_content: dict) -> str:
    """
    PO names live in a rich-text slate like:
      {type: slate, root: {children: [
         {children: [{text: 'Program Outcome: <NAME>', bold: true}]},
         {children: [{text: '<description>'}]}
      ]}}
    Strip the 'Program Outcome: ' prefix and return just the name.

    Preferred alternative: use the .name field from any ref that points at
    this outcome — e.g. alignment-table refs already contain the plain name.
    """
    if not isinstance(slate_content, dict):
        return ""
    root = slate_content.get("root", {})
    first_line = root.get("children", [{}])[0] if root.get("children") else {}
    text = ""
    for span in first_line.get("children", []):
        text += span.get("text", "")
    prefix = "Program Outcome: "
    if text.startswith(prefix):
        return text[len(prefix):].strip()
    return text.strip()


def extract_po_description_from_slate(slate_content: dict) -> str:
    """
    PO descriptions are all lines in the slate AFTER the 'Program Outcome: <n>' title line.
    Returns the flattened body text, or '' if there's no body.
    """
    if not isinstance(slate_content, dict):
        return ""
    root = slate_content.get("root", {})
    lines = root.get("children", [])
    body_parts: list[str] = []
    for line in lines[1:]:
        for span in line.get("children", []):
            body_parts.append(span.get("text", ""))
        body_parts.append("\n")
    return "".join(body_parts).strip()
