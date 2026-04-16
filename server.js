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
const COURSE_TABLE = "grid-8i2Q6-eoTP";
const COMP_TABLE   = "grid-VZwiNNkP1B";
const CCT_TABLE    = "grid-bdbEJT9Kuq";   // _PC | _CCTs
const PO_TABLE     = "grid-tUkDsBwjVg";   // _Progs | _Courses to _Program Outcomes

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
  cctRowIds:     "c-6ucKx1qPTM",  // _PC | _CCTs (linked, already filtered to aligned)
  poRowIds:      "c-jkdSmFgPaR",  // _Progs|_Courses to _Program Outcomes
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

// ─── Value extractors (MCP wraps values in { content: ... }) ─────────────────

function unwrap(val) {
  if (val && typeof val === "object" && "content" in val) return val.content;
  return val;
}

function asString(raw) {
  const v = unwrap(raw);
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object" && v.name) return String(v.name);
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
    const regexSafe = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const filterFormula =
      `[Current Course Code].ToText().RegexMatch("${regexSafe}") OR ` +
      `[Course Name].Contains("${searchTerm}")`;

    // ── Step 1: course row ──────────────────────────────────────────────────
    console.log("  1/2  course row…  filter:", filterFormula);
    const courseResult = await mcpCallTool("table_rows_read", {
      uri: `coda://docs/${docId}/tables/${COURSE_TABLE}`,
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

    // ── Step 2: competencies, CCTs, POs (in parallel) ───────────────────────
    const compIds = asRelationIds(cv[CC.compRowIds]);
    const cctIds  = asRelationIds(cv[CC.cctRowIds]);
    const poIds   = asRelationIds(cv[CC.poRowIds]);
    console.log(`  2/2  ${compIds.length} comps · ${cctIds.length} CCTs · ${poIds.length} POs…`);

    const [compRes, cctRes, poRes] = await Promise.all([
      compIds.length
        ? mcpCallTool("table_rows_read", {
            uri: `coda://docs/${docId}/tables/${COMP_TABLE}`,
            rowNumbersOrIds: compIds,
            rowLimit: 100,
          })
        : Promise.resolve({ rows: [] }),
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
