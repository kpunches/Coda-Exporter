// server.js — Coda CCW Proxy
// Fetches course data from Coda REST API directly (no model in the loop).
// The artifact POSTs a course code here; this server calls Coda, normalizes
// the response, and returns clean JSON the artifact can feed straight into
// its DOCX builder.
//
// Usage:
//   1. cp .env.example .env   # add your Coda API token
//   2. npm install
//   3. npm start              # → http://localhost:3001
//
// The artifact (or any client) calls:
//   POST /api/course  { docId?, courseCode?, courseName? }
//   → returns the structured JSON the DOCX builder expects

require("dotenv").config();
const express = require("express");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json());

// ─── Config ──────────────────────────────────────────────────────────────────

const CODA_API = "https://coda.io/apis/v1";
const CODA_TOKEN = process.env.CODA_API_TOKEN;

if (!CODA_TOKEN) {
  console.error("ERROR: CODA_API_TOKEN is not set. Copy .env.example to .env and add your token.");
  process.exit(1);
}

// Table IDs in the Design & Development V3 doc
const COURSE_TABLE = "grid-8i2Q6-eoTP";
const COMP_TABLE   = "grid-VZwiNNkP1B";

// Course-level column IDs
const CC = {
  courseCode:  "c-sgyJdn2bVc",   // "Current Course Code"
  courseName:  "c-nXrSiX6Q7R",   // "Course Name"
  modality:    "c-sWmJbOaqEx",   // select → .name
  status:      "c-HDAq5esByQ",   // select → .name
  creditUnits: "c-8oRBhcidnK",   // number
  scopeNotes:  "c-fBcr85YOgL",   // slate or plain text
  asmtRationale: "c-k928L_ucpO", // plain text
  evidence:    "c-XLbFh1-a65",   // slate or plain text
  lrStrategy:  "c-lTxHu4kAH6",  // slate or plain text
  tools:       "c-mrpeRPfccr",   // slate or plain text
  compRowIds:  "c-g0MhMSRON7",   // relation array → row identifiers
};

// Competency-level column IDs
const KC = {
  order:      "c-lSdUQo1MHL",   // number
  titleRaw:   "c-uOF-EO1Y7l",   // slate or plain string
  level:      "c-kEEJh01EC8",   // select → .name
  modality:   "c-YfHIST3JLm",   // select → .name
  rationale:  "c-kTJHnj27ap",   // plain text
  evidence:   "c-wQHNTJK75J",   // plain text
  scopeNotes: "c-HEZKT1VZvV",   // plain text
  standards:  "c-TeYyPpbGcv",   // slate or empty
  skills:     "c-yB3khbbLqU",   // relation array → .name = "Cat\nTitle\nStmt"
};

// ─── Value extractors ────────────────────────────────────────────────────────
// Handle both REST API format (valueFormat=rich) and MCP format ({content:…})

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
  // MCP format: {type:"arr", value:[{identifier:"i-xxx"}]}
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value)) {
    return v.value.map(r => r.identifier || r.id || "").filter(Boolean);
  }
  // REST rich format: array of objects
  if (Array.isArray(v)) {
    return v.map(r => (typeof r === "object" ? r.identifier || r.id || "" : String(r))).filter(Boolean);
  }
  // REST simple format: comma-separated string (can't extract IDs — fallback)
  if (typeof v === "string" && v.includes(",")) {
    return v.split(",").map(s => s.trim()).filter(Boolean);
  }
  return [];
}

function asSkills(raw) {
  const v = unwrap(raw);
  if (!v) return [];
  // MCP format: {type:"arr", value:[{name:"Cat\nTitle\nStmt"}]}
  if (typeof v === "object" && v.type === "arr" && Array.isArray(v.value)) {
    return v.value.map(r => ({ name: r.name || "" }));
  }
  // REST rich format: array of objects
  if (Array.isArray(v)) {
    return v.map(r => ({ name: typeof r === "object" ? r.name || "" : String(r) }));
  }
  return [];
}

// ─── Coda REST API caller ────────────────────────────────────────────────────

async function codaGet(path, params = {}) {
  const url = new URL(`${CODA_API}${path}`);
  for (const [k, v] of Object.entries(params)) {
    if (v != null) url.searchParams.set(k, String(v));
  }

  const resp = await fetch(url.toString(), {
    headers: { Authorization: `Bearer ${CODA_TOKEN}` },
  });

  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Coda API ${resp.status}: ${body.slice(0, 500)}`);
  }

  return resp.json();
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

    // Build filter — matches the formula the original artifact used
    const query = courseCode
      ? `[Course Code].Contains("${courseCode}")`
      : `[Course Name].Contains("${courseName}")`;

    // ── Step 1: course row ──────────────────────────────────────────────────
    console.log("  1/2  course row…");
    const courseResp = await codaGet(
      `/docs/${docId}/tables/${COURSE_TABLE}/rows`,
      { query, valueFormat: "rich", limit: 1 }
    );

    const courseRow = (courseResp.items || [])[0];
    if (!courseRow) {
      return res.status(404).json({ error: `Course not found: ${search}` });
    }

    const cv = courseRow.values || {};
    const code = asString(cv[CC.courseCode]).trim();
    const name = asString(cv[CC.courseName]).trim();
    console.log(`       → ${code} — ${name}`);

    // ── Step 2: competency rows (parallel) ──────────────────────────────────
    const compIds = asRelationIds(cv[CC.compRowIds]);
    console.log(`  2/2  ${compIds.length} competency rows…`);

    const competencies = [];
    if (compIds.length) {
      const results = await Promise.allSettled(
        compIds.map(id =>
          codaGet(`/docs/${docId}/tables/${COMP_TABLE}/rows/${id}`, {
            valueFormat: "rich",
          })
        )
      );

      for (const r of results) {
        if (r.status !== "fulfilled") {
          console.warn(`       ⚠ comp fetch failed: ${r.reason?.message}`);
          continue;
        }
        const kv = r.value.values || {};
        competencies.push({
          order:      asNumber(kv[KC.order]),
          titleRaw:   asSlate(kv[KC.titleRaw]),
          level:      asName(kv[KC.level]),
          modality:   asName(kv[KC.modality]),
          rationale:  asString(kv[KC.rationale]),
          evidence:   asString(kv[KC.evidence]),
          scopeNotes: asString(kv[KC.scopeNotes]),
          standards:  asSlate(kv[KC.standards]) || null,
          skills:     asSkills(kv[KC.skills]),
        });
      }
    }

    console.log(`  done ${Date.now() - t0}ms — ${competencies.length} comps`);

    // ── Return exactly the shape the DOCX builder expects ───────────────────
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
      competencies,
    });
  } catch (err) {
    console.error(`  ERROR: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

// ─── Health check ────────────────────────────────────────────────────────────

app.get("/health", (_req, res) => {
  res.json({ ok: true, timestamp: new Date().toISOString() });
});

// ─── Start ───────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n  Coda CCW proxy · http://localhost:${PORT}`);
  console.log(`  POST /api/course   { docId?, courseCode?, courseName? }`);
  console.log(`  GET  /health\n`);
});
