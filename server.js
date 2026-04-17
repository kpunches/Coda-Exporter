// server.js — Coda CCW Proxy (MCP version)
// Connects to Coda's MCP server directly (no model in the loop).
// MCP supports full CFL formulas, which the REST API doesn't.
//
// Usage:
//   1. cp .env.example .env   # add your Coda API token
//   2. npm install
//   3. npm start              # → http://localhost:3001
//
// Client calls:
//   POST /api/course  { docId?, courseCode?, courseName? }

import "dotenv/config";
import express from "express";
import cors from "cors";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";

const app = express();
app.use(cors());
app.use(express.json());

// ─── Config ──────────────────────────────────────────────────────────────────

const CODA_MCP_URL = "https://coda.io/apis/mcp";
const CODA_TOKEN = process.env.CODA_API_TOKEN;

if (!CODA_TOKEN) {
  console.error("ERROR: CODA_API_TOKEN is not set.");
  process.exit(1);
}

// Table IDs in the Design & Development V3 doc
const COURSE_TABLE   = "grid-8i2Q6-eoTP";
const PC_TABLE       = "grid-fEebuvIQBl";   // _Progs | _Courses (program–course pairing)
const COMP_TABLE     = "grid-VZwiNNkP1B";
const CCT_TABLE      = "grid-bdbEJT9Kuq";   // _PC | _CCTs
const PO_TABLE       = "grid-tUkDsBwjVg";   // _Progs | _Courses to _Program Outcomes

// Additional tables used by PDOW (Program Map) export
const PROGRAM_OUTCOMES_TABLE = "grid-aFulSXgnJe"; // _Progs Outcomes
const CCTS_TABLE             = "grid-8Azxi66Mj-"; // _CCTs (base)
const PCC_TABLE              = "grid-TmGA_WNe3_"; // _Progs_Courses | _Comps
const COURSE_X_PO_TABLE      = "grid-tUkDsBwjVg"; // course × PO junction (same as PO_TABLE)
const COURSE_X_CCT_TABLE     = "grid-bdbEJT9Kuq"; // course × CCT junction (same as CCT_TABLE)
const COMP_X_PO_TABLE        = "grid-NkK5JDeDhF"; // competency × PO junction
const COMP_X_CCT_TABLE       = "grid-6xLuZLsQ_t"; // competency × CCT junction

// Course-level column IDs
const CC = {
  courseCode:    "c-sgyJdn2bVc",
  courseName:    "c-nXrSiX6Q7R",
  modality:      "c-sWmJbOaqEx",
  status:        "c-HDAq5esByQ",
  creditUnits:   "c-8oRBhcidnK",
  scopeNotes:    "c-fBcr85YOgL",
  asmtRationale: "c-k928L_ucpO",
  evidence:      "c-XLbFh1-a65",
  lrStrategy:    "c-lTxHu4kAH6",
  tools:         "c-mrpeRPfccr",
  compRowIds:    "c-g0MhMSRON7",
  pcPairingIds:  "c-RaxZ4eQJFf",  // refs to PC_TABLE — the bridge to CCTs/POs
};

// Program–Course pairing row column IDs (PC_TABLE)
const PC = {
  programRef:  "c-UaE_k1Ivfh",     // ref → programs
  courseRef:   "c-aGpuFk4ifn",     // ref → courses (base)
  cctRowIds:   "c-6ucKx1qPTM",     // refs to CCT_TABLE (only aligned ones) — CCW path
  poRowIds:    "c-jkdSmFgPaR",     // refs to PO_TABLE — CCW path
  // PDOW metadata — course-level attributes surfaced in the Program Map xlsx
  stdPath:     "c-tJdYNtbcRy",     // number
  term:        "c-j0Y6lwi_yU",     // text
  cu:          "c-DTsXuT3z72",     // number
  scope:       "c-UlpI6zHbtU",     // ref → select (New/Existing/Redesign)
  ownership:   "c-SSH1vm1eG5",     // ref → schools
  designation: "c-YZLH4IcD51",     // text (Core / Major / GE-…)
  certificate: "c-sj2stjNrbN",     // text (usually blank)
  scopeNotes:  "c-nHeKFgF3C5",     // slate — course-level scope notes
  description: "c-PoBDbaSL6S",     // plain text — course description (Existing)
};

// Competency-level column IDs
const KC = {
  order:      "c-lSdUQo1MHL",
  titleRaw:   "c-uOF-EO1Y7l",
  level:      "c-kEEJh01EC8",
  modality:   "c-YfHIST3JLm",
  rationale:  "c-kTJHnj27ap",
  evidence:   "c-wQHNTJK75J",
  scopeNotes: "c-HEZKT1VZvV",
  standards:  "c-TeYyPpbGcv",
  skills:     "c-yB3khbbLqU",
};

// CCT & PO row column IDs (both tables share this schema)
const TCT = {
  theme:       "c-4CuABuCpCx",   // ref → .name
  aligned:     "c-bBVinldWlI",   // arr of ref → .value[0].name (e.g. "M", "A", "R")
  description: "c-6nEfa4CN4i",   // plain text (CCT only)
};

// _Progs Outcomes (program-outcomes base) columns
const PO_COLS = {
  nameSlate: "c-h3bw7zZXzf",     // slate: "Program Outcome: <name>\n<description>"
  programRef: "c-joP6IMYsmR",    // ref → programs
};

// _CCTs (base) columns
const CCT_COLS = {
  name:        "c-eJu3miLkC_",   // plain text
  description: "c-n4q71ORRVT",   // plain text
  programRef:  "c-joP6IMYsmR",   // ref → programs (same col ID as PO_COLS.programRef by coincidence)
};

// _Progs_Courses | _Comps (PCC) columns — competency-level metadata
const PCC_COLS = {
  programRef:    "c-a4-JBknQCD", // ref → programs
  courseBaseRef: "c-UfDWlbwQXC", // ref → courses (base)
  progCourseRef: "c-WLMD7Ir-os", // ref → prog_courses (junction)
  title:         "c-XRPTQDN2ZN", // plain text
  statement:     "c-ZQ354aYNfj", // plain text
  level:         "c-FZuaohTac9", // number
  assessment:    "c-GunVQKG_Ir", // ref → assessments (name = "PA" / "OA" / etc.)
};

// Course × PO and Course × CCT junction columns (shared schema)
const CXA_COLS = {
  progCourseRef: "c-Dn7elyFRLj", // ref → prog_courses
  target:        "c-4CuABuCpCx", // ref → PO or CCT
  aligned:       "c-bBVinldWlI", // arr of refs to IRMA letters
};

// Competency × PO and Competency × CCT junction columns (shared schema)
const CMX_COLS = {
  pccRef:  "c-pvQqegHRGa",       // ref → PCC
  target:  "c-K0OcZXEwDe",       // ref → PO or CCT
  aligned: "c-BKtOyobxvu",       // arr of refs to IRMA letters
};

// _Programs columns
const PROGRAMS_NAME_TEXT    = "c-JbFOWBC9i0"; // plain-text program name
const PROGRAMS_NAME_DISPLAY = "c-L3c8jGheXt"; // "B.S. ..." / "M.S. ..." display form

// ─── Value extractors (MCP wraps values in { content: ... }) ─────────────────

function unwrap(val) {
  if (val && typeof val === "object" && "content" in val) return val.content;
  return val;
}

function asString(raw) {
  const v = unwrap(raw);
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object") {
    // Ref shape: { name, identifier } — the name is the display text.
    if (v.name) return String(v.name);
    // Slate shape: { root: { children: [...] } } — flatten to plain text so
    // we don't leak "[object Object]" through when a "plain text" column is
    // actually returned as an empty slate by Coda (seen on scopeNotes).
    if (v.root && Array.isArray(v.root.children)) {
      return v.root.children
        .flatMap(line => (line.children || []).map(c => c.text || ""))
        .join("")
        .trim();
    }
    // Unknown / empty object → treat as empty rather than stringifying.
    return "";
  }
  return String(v);
}

function asNumber(raw) {
  const v = unwrap(raw);
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && v.type === "num") return v.value || 0;
  return Number(v) || 0;
}

function asName(raw) {
  const v = unwrap(raw);
  if (typeof v === "string") return v;
  if (v && typeof v === "object" && v.name) return v.name;
  return "";
}

function asSlate(raw) {
  const v = unwrap(raw);
  if (!v) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object" && (v.type === "slate" || v.root)) return v;
  return "";
}

// Single-ref column extractor → returns the refd row identifier ("i-…") or null.
// Covers: direct ref objects ({identifier, name}), MCP-wrapped content,
// and arr-of-one-ref shapes. Returns null if empty.
function extractRefId(raw) {
  const v = unwrap(raw);
  if (!v) return null;
  if (typeof v === "string") return v;
  if (typeof v === "object") {
    if (v.type === "arr" && Array.isArray(v.value) && v.value[0]) {
      return v.value[0].identifier || v.value[0].id || null;
    }
    return v.identifier || v.id || null;
  }
  return null;
}

function asRelationIds(raw) {
  const v = unwrap(raw);
  if (!v) return [];
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value)) {
    return v.value.map(r => r.identifier || r.id || "").filter(Boolean);
  }
  if (Array.isArray(v)) {
    return v.map(r => (typeof r === "object" ? r.identifier || r.id || "" : String(r))).filter(Boolean);
  }
  return [];
}

function asSkills(raw) {
  const v = unwrap(raw);
  if (!v) return [];
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value)) {
    return v.value.map(r => ({ name: r.name || "" }));
  }
  if (Array.isArray(v)) {
    return v.map(r => ({ name: typeof r === "object" ? r.name || "" : String(r) }));
  }
  return [];
}

// Alignment field: can be "" (not aligned), or an arr of ref with .name = "M"/"A"/"R"/etc.
function asAligned(raw) {
  const v = unwrap(raw);
  if (!v) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value) && v.value[0]) {
    return v.value[0].name || "";
  }
  if (typeof v === "object" && v.name) return v.name;
  return "";
}

// ─── Coda MCP caller ─────────────────────────────────────────────────────────
// Opens a fresh MCP session per call. Simple, reliable — re-use isn't worth
// the session-management complexity for this low-volume endpoint.

async function mcpCallTool(toolName, args) {
  const transport = new StreamableHTTPClientTransport(
    new URL(CODA_MCP_URL),
    {
      requestInit: {
        headers: { Authorization: `Bearer ${CODA_TOKEN}` },
      },
    }
  );

  const client = new Client(
    { name: "coda-ccw-proxy", version: "1.0.0" },
    { capabilities: {} }
  );

  try {
    await client.connect(transport);
    const result = await client.callTool({ name: toolName, arguments: args });

    // MCP returns content as [{ type: "text", text: "<json-or-error>" }]
    const text = result?.content?.[0]?.text;
    if (!text) throw new Error(`Empty MCP response from ${toolName}`);

    // If the tool reported an error, surface the error text directly —
    // it's not valid JSON. Check isError flag AND look for common error prefixes.
    if (result.isError || /^(Tool |Error |Exception)/i.test(text)) {
      throw new Error(`MCP ${toolName}: ${text.slice(0, 500)}`);
    }

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (parseErr) {
      throw new Error(`MCP ${toolName} returned non-JSON: ${text.slice(0, 500)}`);
    }

    if (parsed.error) throw new Error(`MCP ${toolName}: ${parsed.error}`);

    // The Coda MCP wraps its result inside { toolName, result: {...} }
    return parsed.result || parsed;
  } finally {
    try { await client.close(); } catch {}
  }
}

// Multi-ref array → list of identifier strings. Used to iterate over refs that
// point at many rows (e.g. an "Aligned" field whose value is [I, R, M]).
function asRefIdArray(raw) {
  const v = unwrap(raw);
  if (!v) return [];
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value)) {
    return v.value.map(r => r.identifier || r.id || "").filter(Boolean);
  }
  if (Array.isArray(v)) return v.map(r => (typeof r === "object" ? r.identifier || r.id || "" : String(r))).filter(Boolean);
  return [];
}

// Multi-ref array → list of ref .name strings. Used for IRMA letter arrays
// (each ref's .name is "I" / "R" / "M" / "A" / "X").
function asRefNames(raw) {
  const v = unwrap(raw);
  if (!v) return [];
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value)) {
    return v.value.map(r => String(r.name || "").trim()).filter(Boolean);
  }
  if (Array.isArray(v)) return v.map(r => (typeof r === "object" ? String(r.name || "").trim() : String(r))).filter(Boolean);
  return [];
}

// Slate shape → plain text (whole body, line-joined). Used for scope notes and
// any other rich-text column we want to store as plain text for the xlsx.
function asSlateText(raw) {
  const v = unwrap(raw);
  if (!v) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object" && v.root && Array.isArray(v.root.children)) {
    return v.root.children
      .map(line => (line.children || []).map(c => c.text || "").join(""))
      .filter(Boolean)
      .join("\n")
      .trim();
  }
  return "";
}

// PO slate parser — matches pdow_exporter/coda_schema.extract_po_name_from_slate.
// The PO name lives on the first line, prefixed with "Program Outcome: "; the
// description is everything after that.
function poNameFromSlate(raw) {
  const v = unwrap(raw);
  if (!v || typeof v !== "object" || !v.root) return "";
  const first = (v.root.children || [])[0] || {};
  const text  = (first.children || []).map(c => c.text || "").join("");
  return text.replace(/^Program Outcome:\s*/, "").trim();
}

function poDescriptionFromSlate(raw) {
  const v = unwrap(raw);
  if (!v || typeof v !== "object" || !v.root) return "";
  const lines = v.root.children || [];
  return lines.slice(1)
    .map(line => (line.children || []).map(c => c.text || "").join(""))
    .join("\n")
    .trim();
}

// ─── Paginated read for tables that may exceed rowLimit (100) ─────────────────
// Coda MCP caps rowLimit at 100 per call; many tables (prog_courses, PCC,
// alignment junctions) regularly exceed that. The MCP doesn't return a page
// token — only `hasMore` + `totalRows` — so we paginate by passing `offset`
// (skip-N-rows) on each subsequent call. 50-iteration safety cap to avoid
// runaway loops if hasMore ever fails to flip.
async function mcpReadAllRows(docId, tableGridId) {
  const uri = `coda://docs/${docId}/tables/${tableGridId}`;
  const allRows = [];
  for (let page = 0; page < 50; page++) {
    const args = { uri, rowLimit: 100 };
    if (allRows.length > 0) args.offset = allRows.length;
    const result = await mcpCallTool("table_rows_read", args);
    if (page === 0) {
      console.log(`  [pagination] ${tableGridId} unfiltered: total=${result.totalRows} hasMore=${result.hasMore} rows=${(result.rows || []).length}`);
    }
    allRows.push(...(result.rows || []));
    if (!result.hasMore) return allRows;
  }
  console.warn(`mcpReadAllRows: hit 50-page safety cap on ${tableGridId}; returning ${allRows.length} rows`);
  return allRows;
}

// Same as mcpReadAllRows but adds a filterFormula and surfaces filterFormulaError
// (trap #1) before the rows are returned. Used for the four PDOW alignment
// junctions — each has a precomputed [Program Abbreviated] / [Program
// Abbreviation] column that makes this cheap.
async function mcpReadAllFilteredRows(docId, tableGridId, filterFormula) {
  const uri = `coda://docs/${docId}/tables/${tableGridId}`;
  const allRows = [];
  for (let page = 0; page < 50; page++) {
    const args = { uri, filterFormula, rowLimit: 100 };
    if (allRows.length > 0) args.offset = allRows.length;
    const result = await mcpCallTool("table_rows_read", args);
    if (result.filterFormulaError) {
      throw new Error(`Coda filter rejected: ${result.filterFormulaError}. Filter: ${filterFormula}`);
    }
    if (page === 0) {
      console.log(`  [pagination] ${tableGridId} filtered: total=${result.totalRows} hasMore=${result.hasMore} rows=${(result.rows || []).length}`);
    }
    allRows.push(...(result.rows || []));
    if (!result.hasMore) return allRows;
  }
  console.warn(`mcpReadAllFilteredRows: hit 50-page cap on ${tableGridId}`);
  return allRows;
}

// ─── CFL escape for user input in filterFormula ──────────────────────────────
// Coda uses double-quoted strings; embedded " and \ must be escaped. Without
// this, a value like `foo"bar` produces a malformed filter that Coda silently
// rejects (→ filterFormulaError populated, rows returned unfiltered — trap #1).
const cflEscape = s => String(s || "")
  .replace(/\\/g, "\\\\")
  .replace(/"/g, '\\"');

// ─── Coda filter traps (from pdow_exporter/RESUME.md) ────────────────────────
// 1. filterFormulaError is SILENT — table_rows_read still returns rows when a
//    filter fails to parse. Always check result.filterFormulaError.
// 2. Use RowId() — .ID() is NOT a CFL function.
// 3. "Table back-reference" columns (e.g. `_Progs | _Courses`) look filterable
//    but aren't scalar; pick a plain text / select / number column instead.
// 4. Column display name can differ across tables even when column IDs match;
//    CFL resolves by display name, not ID. Confirm per-table via document_read.

// ─── Filtered table read with filter-error detection ─────────────────────────
// Wraps mcpCallTool for table_rows_read calls that use filterFormula, so trap
// #1 can't silently return unfiltered rows. For rowNumbersOrIds reads (no
// filter), keep using mcpCallTool directly — there's no filter to validate.
async function readFilteredTable({ docId, tableGridId, filterFormula, rowLimit }) {
  const result = await mcpCallTool("table_rows_read", {
    uri: `coda://docs/${docId}/tables/${tableGridId}`,
    filterFormula,
    rowLimit,
  });

  if (result.filterFormulaError) {
    throw new Error(`Coda filter rejected: ${result.filterFormulaError}. Filter: ${filterFormula}`);
  }

  return result;
}

// ─── API endpoint ────────────────────────────────────────────────────────────

app.post("/api/course", async (req, res) => {
  const t0 = Date.now();
  try {
    const { docId = "4YIajnJqvo", courseCode, courseName } = req.body;

    if (!courseCode && !courseName) {
      return res.status(400).json({ error: "courseCode or courseName is required" });
    }

    const search = courseCode || courseName;
    console.log(`\n[${new Date().toISOString()}] Fetching "${search}" from doc ${docId}`);

    // CFL formula — check BOTH columns so it works regardless of which
    // field the user typed the value into. RegexMatch on ToText() is
    // whitespace- and case-forgiving for course codes.
    const searchTerm = String(search).trim();
    // Escape regex metacharacters so values with dots, parens etc. still work
    // inside the RegexMatch pattern. cflEscape then handles CFL string quoting
    // for both halves — same rawSearchTerm, different escape layers.
    const regexSafe = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const filterFormula =
      `[Current Course Code].ToText().RegexMatch("${cflEscape(regexSafe)}") OR ` +
      `[Course Name].Contains("${cflEscape(searchTerm)}")`;

    // ── Step 1: course row ──────────────────────────────────────────────────
    console.log("  1/2  course row…  filter:", filterFormula);
    const courseResult = await readFilteredTable({
      docId,
      tableGridId: COURSE_TABLE,
      filterFormula,
      rowLimit: 1,
    });

    const courseRow = (courseResult.rows || [])[0];
    if (!courseRow) {
      return res.status(404).json({ error: `Course not found: ${search}` });
    }

    const cv = courseRow.values || {};
    const code = asString(cv[CC.courseCode]).trim();
    const name = asString(cv[CC.courseName]).trim();
    console.log(`       → ${code} — ${name}`);

    // ── Step 2: program-course pairing (bridge to CCTs/POs) + competencies ──
    const compIds       = asRelationIds(cv[CC.compRowIds]);
    const pcPairingIds  = asRelationIds(cv[CC.pcPairingIds]);
    console.log(`  2/3  ${compIds.length} comps + ${pcPairingIds.length} PC-pairing rows…`);

    const [compRes, pcRes] = await Promise.all([
      compIds.length
        ? mcpCallTool("table_rows_read", {
            uri: `coda://docs/${docId}/tables/${COMP_TABLE}`,
            rowNumbersOrIds: compIds,
            rowLimit: 100,
          })
        : Promise.resolve({ rows: [] }),
      pcPairingIds.length
        ? mcpCallTool("table_rows_read", {
            uri: `coda://docs/${docId}/tables/${PC_TABLE}`,
            rowNumbersOrIds: pcPairingIds,
            rowLimit: 10,
          })
        : Promise.resolve({ rows: [] }),
    ]);

    // From the first program-course pairing row, get linked CCT and PO row IDs
    const pcRow = (pcRes.rows || [])[0];
    const pcv   = pcRow?.values || {};
    const cctIds = asRelationIds(pcv[PC.cctRowIds]);
    const poIds  = asRelationIds(pcv[PC.poRowIds]);

    // ── Step 3: CCTs and POs (parallel) ─────────────────────────────────────
    console.log(`  3/3  ${cctIds.length} CCTs · ${poIds.length} POs…`);
    const [cctRes, poRes] = await Promise.all([
      cctIds.length
        ? mcpCallTool("table_rows_read", {
            uri: `coda://docs/${docId}/tables/${CCT_TABLE}`,
            rowNumbersOrIds: cctIds,
            rowLimit: 100,
          })
        : Promise.resolve({ rows: [] }),
      poIds.length
        ? mcpCallTool("table_rows_read", {
            uri: `coda://docs/${docId}/tables/${PO_TABLE}`,
            rowNumbersOrIds: poIds,
            rowLimit: 100,
          })
        : Promise.resolve({ rows: [] }),
    ]);

    const competencies = (compRes.rows || []).map(row => {
      const kv = row.values || {};
      return {
        order:      asNumber(kv[KC.order]),
        titleRaw:   asSlate(kv[KC.titleRaw]),
        level:      asName(kv[KC.level]),
        modality:   asName(kv[KC.modality]),
        rationale:  asString(kv[KC.rationale]),
        evidence:   asString(kv[KC.evidence]),
        scopeNotes: asString(kv[KC.scopeNotes]),
        standards:  asSlate(kv[KC.standards]) || null,
        skills:     asSkills(kv[KC.skills]),
      };
    });

    const ccts = (cctRes.rows || []).map(row => {
      const kv = row.values || {};
      return {
        name:        asName(kv[TCT.theme]),
        description: asString(kv[TCT.description]),
        aligned:     asAligned(kv[TCT.aligned]),
      };
    }).filter(c => c.name);

    const pos = (poRes.rows || []).map(row => {
      const kv = row.values || {};
      return {
        name:    asName(kv[TCT.theme]),
        aligned: asAligned(kv[TCT.aligned]),
      };
    }).filter(p => p.name);

    console.log(`  done ${Date.now() - t0}ms — ${competencies.length} comps · ${ccts.length} CCTs · ${pos.length} POs`);

    res.json({
      courseCode:                  code,
      courseName:                  name,
      modality:                    asName(cv[CC.modality]),
      status:                      asName(cv[CC.status]),
      creditUnits:                 asNumber(cv[CC.creditUnits]),
      scopeNotes:                  asSlate(cv[CC.scopeNotes]),
      assessmentModalityRationale: asString(cv[CC.asmtRationale]),
      evidence:                    asSlate(cv[CC.evidence]),
      lrStrategy:                  asSlate(cv[CC.lrStrategy]),
      tools:                       asSlate(cv[CC.tools]),
      ccts,
      pos,
      competencies,
    });
  } catch (err) {
    console.error(`  ERROR: ${err.message}`);
    console.error(err.stack);
    res.status(500).json({ error: err.message });
  }
});

// ─── Programs (PDOW) ─────────────────────────────────────────────────────────
// Returns the program list for the dropdown. Each program: {rowId, abbr, name}.
// abbr (MSCSIA, MSCIN, etc.) is the canonical key used by the four junction
// tables and by every other PDOW endpoint, so the frontend keys its state on
// it rather than rowId.

const PROGRAMS_TABLE = "grid-_cLrawcUzd";
const PROGRAMS_ABBR  = "c-WKn3oxZN4y"; // "Program Abbreviation" — plain text

app.get("/api/programs", async (req, res) => {
  const t0 = Date.now();
  try {
    const docId = req.query.docId || "4YIajnJqvo";
    console.log(`\n[${new Date().toISOString()}] GET /api/programs docId=${docId}`);

    // No filterFormula — we want every program row. Calling mcpCallTool
    // directly since readFilteredTable assumes a filter is present.
    // rowLimit cap is 100 on the Coda MCP. If WGU ever exceeds 100 programs,
    // switch to paginated reads via pageToken.
    const result = await mcpCallTool("table_rows_read", {
      uri: `coda://docs/${docId}/tables/${PROGRAMS_TABLE}`,
      rowLimit: 100,
    });

    const programs = (result.rows || [])
      .map(r => {
        const v = r.values || {};
        return {
          rowId: r.id || r.rowId,
          abbr:  asString(v[PROGRAMS_ABBR]).trim(),
          name:  asString(v["c-L3c8jGheXt"]).trim()        // PROGRAMS_DISPLAY ("M.S. ...")
                 || asString(v["c-JbFOWBC9i0"]).trim(),    // PROGRAMS_NAME_TEXT (fallback)
        };
      })
      .filter(p => p.abbr)                                  // drop rows without an abbreviation
      .sort((a, b) => a.abbr.localeCompare(b.abbr));        // dropdown order

    console.log(`  done ${Date.now() - t0}ms — ${programs.length} programs`);
    res.json({ programs });
  } catch (err) {
    console.error(`  ERROR: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

// ─── Courses for a program (PDOW + future course dropdown) ──────────────────
// Given a program abbreviation (MSCSIA, MSCIN, …), return the courses
// belonging to that program as [{progCourseId, code, name}] sorted by code.
//
// Strategy (matches pdow_exporter/extractor.py):
//   1. Fetch _Programs, find the row whose abbreviation matches → programRowId.
//   2. Fetch _Progs | _Courses (paginated) — junction of program × course.
//   3. Client-side filter to rows whose program ref matches programRowId.
//      (Filtering a ref column server-side with CFL is fragile — see RESUME
//      trap #3. Reading and filtering here is cheap and reliable.)
//   4. Join against _Courses (base) to pull code + name for each match.
app.get("/api/courses", async (req, res) => {
  const t0 = Date.now();
  try {
    const docId = req.query.docId || "4YIajnJqvo";
    const programAbbr = String(req.query.programAbbr || "").trim();
    if (!programAbbr) {
      return res.status(400).json({ error: "programAbbr is required" });
    }

    console.log(`\n[${new Date().toISOString()}] GET /api/courses programAbbr=${programAbbr} docId=${docId}`);

    // 1. Resolve the program row ID from the abbreviation.
    const progResult = await mcpCallTool("table_rows_read", {
      uri: `coda://docs/${docId}/tables/${PROGRAMS_TABLE}`,
      rowLimit: 100,
    });
    const programRow = (progResult.rows || []).find(r =>
      asString((r.values || {})[PROGRAMS_ABBR]).trim() === programAbbr
    );
    if (!programRow) {
      return res.status(404).json({ error: `Program not found: ${programAbbr}` });
    }
    const programRowId = programRow.id || programRow.rowId;
    console.log(`  program ${programAbbr} → ${programRowId}`);

    // 2 + 3. Paginated fetch of prog_courses, filter by program ref.
    const allPcRows = await mcpReadAllRows(docId, PC_TABLE);
    const matchedPcRows = allPcRows.filter(r =>
      extractRefId((r.values || {})[PC.programRef]) === programRowId
    );
    console.log(`  prog_courses: ${matchedPcRows.length}/${allPcRows.length} match`);

    // 4. Build a course-base lookup so we can join for code + name.
    const allBaseRows = await mcpReadAllRows(docId, COURSE_TABLE);
    const baseByRowId = {};
    for (const r of allBaseRows) {
      baseByRowId[r.id || r.rowId] = {
        code: asString((r.values || {})[CC.courseCode]).trim(),
        name: asString((r.values || {})[CC.courseName]).trim(),
      };
    }

    const courses = matchedPcRows
      .map(r => {
        const baseId = extractRefId((r.values || {})[PC.courseRef]);
        const base = baseByRowId[baseId] || {};
        return {
          progCourseId: r.id || r.rowId,
          code: base.code || "",
          name: base.name || "",
        };
      })
      .filter(c => c.code)                                        // drop rows with no code
      .sort((a, b) => a.code.localeCompare(b.code));

    console.log(`  done ${Date.now() - t0}ms — ${courses.length} courses`);
    res.json({ courses });
  } catch (err) {
    console.error(`  ERROR: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

// ─── PDOW Program Map data bundle ────────────────────────────────────────────
// Builds the data.json shape consumed by pdow_kit/merger.py: the base model
// (program, POs, CCTs, courses with competencies) plus the 4 raw alignment
// arrays. Frontend zips this with the kit assets for local xlsx build.
//
// Fetches in parallel where possible (9 total):
//   programs, prog_courses, courses base, PCC, POs, CCTs, 4 alignment junctions
// Base tables (PO, CCT, PCC, prog_courses) are filtered client-side by program
// ref — CFL ref filtering is unreliable (RESUME trap #3). Alignment junctions
// use the precomputed [Program Abbreviat(ed|ion)] plain-text column for
// cheap server-side filtering (RESUME § Filter breakthrough).
app.get("/api/pdow-data", async (req, res) => {
  const t0 = Date.now();
  try {
    const docId = req.query.docId || "4YIajnJqvo";
    const programAbbr = String(req.query.programAbbr || "").trim();
    if (!programAbbr) {
      return res.status(400).json({ error: "programAbbr is required" });
    }

    console.log(`\n[${new Date().toISOString()}] GET /api/pdow-data programAbbr=${programAbbr}`);

    // 1. Resolve program row ID (sequential — everything else needs it)
    const progResult = await mcpCallTool("table_rows_read", {
      uri: `coda://docs/${docId}/tables/${PROGRAMS_TABLE}`,
      rowLimit: 100,
    });
    const programRow = (progResult.rows || []).find(r =>
      asString((r.values || {})[PROGRAMS_ABBR]).trim() === programAbbr
    );
    if (!programRow) {
      return res.status(404).json({ error: `Program not found: ${programAbbr}` });
    }
    const programRowId = programRow.id || programRow.rowId;
    const programName  = asString(programRow.values[PROGRAMS_NAME_DISPLAY]).trim()
                      || asString(programRow.values[PROGRAMS_NAME_TEXT]).trim();
    console.log(`  program ${programAbbr} → ${programRowId}`);

    // 2. Parallel fetches (8 reads, all paginated where needed)
    const junctionFilter1 = `[Program Abbreviated] = "${cflEscape(programAbbr)}"`;  // course-level
    const junctionFilter2 = `[Program Abbreviation] = "${cflEscape(programAbbr)}"`; // comp-level

    const [
      pcRows, baseCourseRows, pccRows, poBaseRows, cctBaseRows,
      coursePoRows, courseCctRows, compPoRows, compCctRows,
    ] = await Promise.all([
      mcpReadAllRows(docId, PC_TABLE),
      mcpReadAllRows(docId, COURSE_TABLE),
      mcpReadAllRows(docId, PCC_TABLE),
      mcpReadAllRows(docId, PROGRAM_OUTCOMES_TABLE),
      mcpReadAllRows(docId, CCTS_TABLE),
      mcpReadAllFilteredRows(docId, COURSE_X_PO_TABLE,  junctionFilter1),
      mcpReadAllFilteredRows(docId, COURSE_X_CCT_TABLE, junctionFilter1),
      mcpReadAllFilteredRows(docId, COMP_X_PO_TABLE,    junctionFilter2),
      mcpReadAllFilteredRows(docId, COMP_X_CCT_TABLE,   junctionFilter2),
    ]);

    console.log(`  fetched: ${pcRows.length} PC · ${baseCourseRows.length} courses · ${pccRows.length} PCC · ${poBaseRows.length} POs · ${cctBaseRows.length} CCTs`);
    console.log(`  junctions: ${coursePoRows.length}/${courseCctRows.length}/${compPoRows.length}/${compCctRows.length} (c×po/c×cct/p×po/p×cct)`);

    // 3. Program outcomes for this program (client-side filter on program ref)
    const program_outcomes = poBaseRows
      .filter(r => extractRefId((r.values || {})[PO_COLS.programRef]) === programRowId)
      .map(r => ({
        id: r.id || r.rowId,
        name: poNameFromSlate((r.values || {})[PO_COLS.nameSlate]),
        description: poDescriptionFromSlate((r.values || {})[PO_COLS.nameSlate]),
      }));

    // 4. CCTs for this program
    const ccts = cctBaseRows
      .filter(r => extractRefId((r.values || {})[CCT_COLS.programRef]) === programRowId)
      .map(r => ({
        id:   r.id || r.rowId,
        name: asString((r.values || {})[CCT_COLS.name]).trim(),
        description: asString((r.values || {})[CCT_COLS.description]).trim(),
      }))
      .filter(c => c.name);

    // 5. Courses — filter prog_courses by program ref, join to base table
    const baseByRowId = {};
    for (const r of baseCourseRows) {
      baseByRowId[r.id || r.rowId] = {
        code: asString((r.values || {})[CC.courseCode]).trim(),
        name: asString((r.values || {})[CC.courseName]).trim(),
        cu:   asNumber((r.values || {})[CC.creditUnits]),
      };
    }

    const matchedPc = pcRows.filter(r =>
      extractRefId((r.values || {})[PC.programRef]) === programRowId
    );

    const coursesByPcId = {};
    const courses = matchedPc.map(r => {
      const v = r.values || {};
      const courseBaseId = extractRefId(v[PC.courseRef]);
      const base = baseByRowId[courseBaseId] || {};
      const termText = asString(v[PC.term]);
      const c = {
        prog_course_id:  r.id || r.rowId,
        course_base_id:  courseBaseId,
        code:            base.code || "",
        name:            base.name || "",
        display_name:    (() => {
          const raw = unwrap(v[PC.courseRef]);
          return (raw && typeof raw === "object" && raw.name) ? String(raw.name) : "";
        })(),
        term:            termText,
        std_path_order:  asNumber(v[PC.stdPath]) || null,
        cu:              asNumber(v[PC.cu]) || base.cu || null,
        scope:           asName(v[PC.scope]),
        ownership:       asName(v[PC.ownership]),
        designation:     asString(v[PC.designation]),
        certificate:     asString(v[PC.certificate]),
        scope_notes:     asSlateText(v[PC.scopeNotes]),
        description:     asString(v[PC.description]),
        po_alignments:   {},
        cct_alignments:  {},
        competencies:    [],
      };
      coursesByPcId[c.prog_course_id] = c;
      return c;
    });

    // Stable sort per extractor.py: (term_int, std_path_order, code)
    courses.sort((a, b) => {
      const ta = /^\d+$/.test(a.term) ? parseInt(a.term, 10) : 99;
      const tb = /^\d+$/.test(b.term) ? parseInt(b.term, 10) : 99;
      if (ta !== tb) return ta - tb;
      const sa = a.std_path_order ?? 9999;
      const sb = b.std_path_order ?? 9999;
      if (sa !== sb) return sa - sb;
      return (a.code || "zzz").localeCompare(b.code || "zzz");
    });

    // 6. Competencies — filter PCC by program ref, attach to courses
    for (const r of pccRows) {
      const v = r.values || {};
      if (extractRefId(v[PCC_COLS.programRef]) !== programRowId) continue;
      let progCourseId = extractRefId(v[PCC_COLS.progCourseRef]);
      if (!progCourseId) {
        // Fallback: resolve via course base ref → matching prog_course for this program
        const courseBaseId = extractRefId(v[PCC_COLS.courseBaseRef]);
        const host = courses.find(c => c.course_base_id === courseBaseId);
        if (host) progCourseId = host.prog_course_id;
      }
      const comp = {
        pcc_id:         r.id || r.rowId,
        prog_course_id: progCourseId || null,
        title:          asString(v[PCC_COLS.title]),
        statement:      asString(v[PCC_COLS.statement]),
        level:          asNumber(v[PCC_COLS.level]) || null,
        assessment:     asName(v[PCC_COLS.assessment]),
        po_alignments:  {},
        cct_alignments: {},
      };
      const host = coursesByPcId[progCourseId];
      if (host) host.competencies.push(comp);
    }

    // 7. Raw alignment arrays — merger fills them into the model at build time
    const course_po = coursePoRows
      .map(r => ({
        prog_course_id: extractRefId((r.values || {})[CXA_COLS.progCourseRef]),
        po_id:          extractRefId((r.values || {})[CXA_COLS.target]),
        letters:        asRefNames((r.values || {})[CXA_COLS.aligned]),
      }))
      .filter(x => x.prog_course_id && x.po_id);

    const course_cct = courseCctRows
      .map(r => ({
        prog_course_id: extractRefId((r.values || {})[CXA_COLS.progCourseRef]),
        cct_id:         extractRefId((r.values || {})[CXA_COLS.target]),
        letters:        asRefNames((r.values || {})[CXA_COLS.aligned]),
      }))
      .filter(x => x.prog_course_id && x.cct_id);

    // For competency-level rows we also need course_base_id since the merger
    // keys on (course_base_id, pcc_id). Derive it from the PCC lookup.
    const pccToCourseBase = {};
    for (const r of pccRows) {
      const pccId = r.id || r.rowId;
      pccToCourseBase[pccId] = extractRefId((r.values || {})[PCC_COLS.courseBaseRef]);
    }

    const comp_po = compPoRows
      .map(r => {
        const pcc = extractRefId((r.values || {})[CMX_COLS.pccRef]);
        return {
          pcc_id:         pcc,
          course_base_id: pccToCourseBase[pcc] || null,
          po_id:          extractRefId((r.values || {})[CMX_COLS.target]),
          letters:        asRefNames((r.values || {})[CMX_COLS.aligned]),
        };
      })
      .filter(x => x.pcc_id && x.po_id);

    const comp_cct = compCctRows
      .map(r => {
        const pcc = extractRefId((r.values || {})[CMX_COLS.pccRef]);
        return {
          pcc_id:         pcc,
          course_base_id: pccToCourseBase[pcc] || null,
          cct_id:         extractRefId((r.values || {})[CMX_COLS.target]),
          letters:        asRefNames((r.values || {})[CMX_COLS.aligned]),
        };
      })
      .filter(x => x.pcc_id && x.cct_id);

    console.log(`  done ${Date.now() - t0}ms — ${courses.length} courses · ${program_outcomes.length} POs · ${ccts.length} CCTs · alignments ${course_po.length}/${course_cct.length}/${comp_po.length}/${comp_cct.length}`);

    res.json({
      program_code: programAbbr,
      model: {
        program: { id: programRowId, name: programName },
        program_outcomes,
        ccts,
        courses,
      },
      alignments: { course_po, course_cct, comp_po, comp_cct },
    });
  } catch (err) {
    console.error(`  ERROR: ${err.message}`);
    console.error(err.stack);
    res.status(500).json({ error: err.message });
  }
});

// ─── Diagnostic: probe table_rows_read schema via intentional validation error
// The Coda MCP schema isn't documented anywhere I can find, and the response
// doesn't include a page token — yet hasMore=true on large tables means some
// pagination mechanism must exist. Send obviously-invalid types for common
// pagination param names and return the Zod error text (which previously told
// us rowLimit maxes at 100). Hit /api/_probe once and the error text in the
// response will reveal the accepted schema.
app.get("/api/_probe", async (req, res) => {
  const tests = [
    { name: "offset_string",   args: { uri: `coda://docs/4YIajnJqvo/tables/${PROGRAMS_TABLE}`, rowLimit: 1, offset: "not-a-number" } },
    { name: "skip_string",     args: { uri: `coda://docs/4YIajnJqvo/tables/${PROGRAMS_TABLE}`, rowLimit: 1, skip: "not-a-number" } },
    { name: "page_string",     args: { uri: `coda://docs/4YIajnJqvo/tables/${PROGRAMS_TABLE}`, rowLimit: 1, page: "not-a-number" } },
    { name: "pageToken_bad",   args: { uri: `coda://docs/4YIajnJqvo/tables/${PROGRAMS_TABLE}`, rowLimit: 1, pageToken: 12345 } },
    { name: "bogus_param",     args: { uri: `coda://docs/4YIajnJqvo/tables/${PROGRAMS_TABLE}`, rowLimit: 1, totallyMadeUp: "xyz" } },
    { name: "rowLimit_1001",   args: { uri: `coda://docs/4YIajnJqvo/tables/${PROGRAMS_TABLE}`, rowLimit: 1001 } },
  ];
  const results = [];
  for (const t of tests) {
    try {
      const r = await mcpCallTool("table_rows_read", t.args);
      results.push({ test: t.name, ok: true, hasMore: r.hasMore, totalRows: r.totalRows, rowCount: (r.rows || []).length });
    } catch (err) {
      results.push({ test: t.name, ok: false, error: String(err.message).slice(0, 1200) });
    }
  }
  res.json({ results });
});

// ─── Health ──────────────────────────────────────────────────────────────────

app.get("/health", (_req, res) => {
  res.json({ ok: true, timestamp: new Date().toISOString() });
});

app.get("/", (_req, res) => {
  res.json({
    name: "Coda CCW Proxy",
    endpoints: ["POST /api/course", "GET /health"],
  });
});

// ─── Start ───────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n  Coda CCW proxy (MCP) · http://localhost:${PORT}`);
  console.log(`  POST /api/course   { docId?, courseCode?, courseName? }`);
  console.log(`  GET  /health\n`);
});
