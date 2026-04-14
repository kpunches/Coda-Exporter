// ─────────────────────────────────────────────────────────────────────────────
// CCW Exporter — Cloudflare Worker
//
// Environment variables (set in Cloudflare dashboard → Worker → Settings → Variables):
//   ANTHROPIC_API_KEY   your Anthropic API key (mark as Secret)
//   CODA_API_TOKEN      your Coda API token    (mark as Secret)
//   ALLOWED_ORIGIN      your GitHub Pages URL, e.g. https://yourname.github.io
//
// Deploy:
//   1. wrangler login
//   2. wrangler deploy
// ─────────────────────────────────────────────────────────────────────────────

export default {
  async fetch(request, env) {

    const origin = request.headers.get("Origin") || "";
    const allowedOrigin = env.ALLOWED_ORIGIN || "*";

    // ── CORS preflight ────────────────────────────────────────────────────────
    if (request.method === "OPTIONS") {
      return new Response(null, {
        headers: {
          "Access-Control-Allow-Origin": allowedOrigin,
          "Access-Control-Allow-Methods": "POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type",
          "Access-Control-Max-Age": "86400",
        },
      });
    }

    const corsHeaders = {
      "Access-Control-Allow-Origin": allowedOrigin,
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    if (request.method !== "POST") {
      return new Response("Method not allowed", { status: 405, headers: corsHeaders });
    }

    // ── Parse request body ────────────────────────────────────────────────────
    let body;
    try {
      body = await request.json();
    } catch {
      return new Response(JSON.stringify({ error: "Invalid JSON body" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const { docId = "4YIajnJqvo", courseCode, courseName, docType = "CCW" } = body;

    if (!courseCode && !courseName) {
      return new Response(JSON.stringify({ error: "Provide courseCode or courseName" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // ── Step 1: Ask Claude to fetch + return structured CCW data via Coda MCP ─
    let courseData;
    try {
      courseData = await fetchCourseDataViaClaude(env, docId, courseCode, courseName);
    } catch (err) {
      return new Response(JSON.stringify({ error: `Data fetch failed: ${err.message}` }), {
        status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    if (courseData.error) {
      return new Response(JSON.stringify({ error: courseData.error }), {
        status: 404, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // ── Step 2: Build the .docx via Claude ────────────────────────────────────
    let docxBase64;
    try {
      docxBase64 = await buildDocxViaClaude(env, courseData, docType);
    } catch (err) {
      return new Response(JSON.stringify({ error: `Document build failed: ${err.message}` }), {
        status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // ── Step 3: Decode base64 → binary and stream back ───────────────────────
    const binary = Uint8Array.from(atob(docxBase64), c => c.charCodeAt(0));
    const filename = `${(courseData.courseCode || "COURSE").replace(/\s/g, "")}_${docType}.docx`;

    return new Response(binary, {
      status: 200,
      headers: {
        ...corsHeaders,
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Content-Length": binary.byteLength.toString(),
      },
    });
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// fetchCourseDataViaClaude
// Calls the Anthropic API with Coda MCP enabled.
// Returns structured JSON with all CCW fields.
// ─────────────────────────────────────────────────────────────────────────────
async function fetchCourseDataViaClaude(env, docId, courseCode, courseName) {
  const searchTerm = courseCode || courseName;
  const filterFormula = courseCode
    ? `[Current Course Code].Contains("${courseCode}")`
    : `[Course Name].Contains("${courseName}")`;

  const systemPrompt = `You are a data extraction assistant. Your job is to read data from a Coda document using the provided MCP tools and return it as structured JSON. You must:
1. Query the _Courses table (grid-8i2Q6-eoTP) in doc ${docId} using the filter: ${filterFormula}
2. Get the matched course row and extract all CCW fields
3. Get all competency rows from _Courses | _Comps (grid-VZwiNNkP1B) using the competency row IDs from field c-g0MhMSRON7
4. Return ONLY a single valid JSON object — no markdown, no explanation, no code fences

The JSON must have this exact shape:
{
  "courseCode": string,
  "courseName": string,
  "modality": string,
  "status": string,
  "creditUnits": number,
  "scopeNotes": object,
  "assessmentModalityRationale": string,
  "evidence": object,
  "lrStrategy": object,
  "tools": object,
  "competencies": [
    {
      "order": number,
      "titleRaw": object or string,
      "level": string,
      "modality": string,
      "rationale": string,
      "evidence": string,
      "scopeNotes": string,
      "standards": object or null,
      "skills": [{ "name": string }]
    }
  ]
}

Field mapping for _Courses row:
- courseCode: c-sgyJdn2bVc
- courseName: c-nXrSiX6Q7R
- modality: c-sWmJbOaqEx (use .name)
- status: c-HDAq5esByQ (use .name)
- creditUnits: c-8oRBhcidnK
- scopeNotes: c-fBcr85YOgL (slate object)
- assessmentModalityRationale: c-k928L_ucpO (plain text)
- evidence: c-XLbFh1-a65 (slate object)
- tools: c-mrpeRPfccr (slate object)
- lrStrategy: c-lTxHu4kAH6 (slate object)
- competencyRowIds: c-g0MhMSRON7 (array of refs, use .value[].identifier)

Field mapping for each _Courses | _Comps row:
- order: c-lSdUQo1MHL
- titleRaw: c-uOF-EO1Y7l (slate or plain string)
- level: c-kEEJh01EC8 (use .name)
- modality: c-YfHIST3JLm (use .name)
- rationale: c-kTJHnj27ap
- evidence: c-wQHNTJK75J
- scopeNotes: c-HEZKT1VZvV
- standards: c-TeYyPpbGcv (slate or null)
- skills: c-yB3khbbLqU (array of refs, use .value[].name — each name is "Category\\nTitle\\nStatement")

If the course is not found, return: { "error": "Course not found: ${searchTerm}" }`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "mcp-client-2025-04-04",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 8000,
      system: systemPrompt,
      messages: [{ role: "user", content: `Fetch CCW data for course: ${searchTerm} from Coda doc ${docId}` }],
      mcp_servers: [
        {
          type: "url",
          url: "https://coda.io/apis/mcp",
          name: "coda",
          authorization_token: env.CODA_API_TOKEN,
        },
      ],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Anthropic API error ${response.status}: ${err}`);
  }

  const data = await response.json();

  // Extract text content from response
  const textBlock = data.content?.find(b => b.type === "text");
  if (!textBlock) throw new Error("No text response from Claude");

  // Strip any accidental markdown fences
  const raw = textBlock.text.replace(/^```(?:json)?\n?/i, "").replace(/\n?```$/i, "").trim();

  try {
    return JSON.parse(raw);
  } catch {
    throw new Error(`Claude returned invalid JSON: ${raw.slice(0, 200)}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// buildDocxViaClaude
// Sends structured course data to Claude and asks it to generate a
// base64-encoded .docx using the docx npm library via a Node.js script.
// Claude returns the script output as base64.
// ─────────────────────────────────────────────────────────────────────────────
async function buildDocxViaClaude(env, courseData, docType) {
  const systemPrompt = `You are a document generation assistant. You will receive structured JSON data for a ${docType} (Course Competency Worksheet) and must generate a valid .docx file using the docx npm library.

Return ONLY a base64-encoded string of the .docx binary — no markdown, no explanation, no code fences, just the raw base64 string.

Document spec:
- Landscape orientation (width:12240, height:15840, orientation:LANDSCAPE)
- Margins: top:1080, right:720, bottom:720, left:720
- Font: Arial, 10pt (size:20 in half-points)
- Content width: 14400 DXA
- Color palette: HEADER_FILL=1F4E79, SUBHEAD_FILL=2E75B6, LIGHT_FILL=D6E4F0, GRID_FILL=F2F7FB
- Page header: "CourseCode – CourseName  |  Page N"

Document structure:
1. Title block: course code + name (bold, dark blue), modality/status/units line
2. Course Level Scope Notes table (full-width, dark header row, content below)
3. Course Level Assessment Strategy table (Modality Rationale row + Evidence row)
4. Course Level LR Strategy table
5. Course Level Tools & Technologies table
6. Section heading: "Course Competencies and Context Tables"
7. One competency table per competency ordered by .order field, each with:
   - Dark header row: "Competency N | Title: ... | Statement: ... | Target Level: ..."
   - Light "Skills" banner row
   - Column headers: Skill Category | Skill Title | Skill Statement (blue header)
   - One row per skill (parse each skill.name by splitting on "\\n" → [category, title, statement])
   - Light "Competency Detail" banner
   - Assessment row: Modality + Rationale + Evidence
   - Scope Notes row
   - Standards Alignment row

Slate JSON parser: slate objects have .root.children[] where each child has .style ("Paragraph"|"BulletedList") and .children[] with {text, bold, italic, url} or InlineStructuredValue with {value:{url,name}}. Convert to docx Paragraphs preserving bold, bullets, and ExternalHyperlinks.

Column widths: label=2200, data=12200; skills: cat=2880, title=4320, statement=7200
All column widths must sum to 14400.
Use ShadingType.CLEAR for all shading.
Use LevelFormat.BULLET with numbering config for bullet lists.
Include Hyperlink character style.

The course data JSON follows.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 8000,
      system: systemPrompt,
      messages: [{
        role: "user",
        content: `Generate the ${docType} Word document for this course data and return base64:\n\n${JSON.stringify(courseData, null, 2)}`
      }],
      tools: [{
        type: "computer_use_20250124",
        name: "computer",
        display_width_px: 1280,
        display_height_px: 800,
      }],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Anthropic API error ${response.status}: ${err}`);
  }

  const data = await response.json();
  const textBlock = data.content?.find(b => b.type === "text");
  if (!textBlock) throw new Error("No base64 response from Claude");

  const raw = textBlock.text.trim();
  // Validate it looks like base64
  if (!/^[A-Za-z0-9+/]+=*$/.test(raw.replace(/\s/g, ""))) {
    throw new Error("Claude did not return valid base64");
  }
  return raw.replace(/\s/g, "");
}
