/**
 * build_ccw.js  —  CCW Word Doc Generator (updated)
 *
 * Changes from previous version:
 *  - Title block: centered, subtitle "Competency Context Worksheet" added
 *  - New section: Course Level Cross Cutting Themes (filtered to non-blank Aligned)
 *  - New section: Course Level Program Outcomes (filtered to non-blank Aligned)
 *  - Both sections placed after Tools & Tech, before Competency tables
 *  - Verbatim data from Coda; blank = omitted section (no placeholders)
 *  - Final QA check: field count logged before writing
 */

'use strict';

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, HeadingLevel, Header,
  PageNumber, NumberFormat, ExternalHyperlink, ShadingType,
  PageOrientation, HeightRule
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── DATA (injected by orchestrator) ─────────────────────────────────────────
const DATA = JSON.parse(fs.readFileSync('/home/claude/ccw_data.json', 'utf8'));

const course = DATA.course;
const comps   = DATA.comps;   // sorted by compOrder
const ccts    = DATA.ccts;    // [{name, aligned, description}] — non-blank aligned only
const pos     = DATA.pos;     // [{name, aligned}] — non-blank aligned only

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const HEADER_FILL   = '1F4E79';
const SUBHEAD_FILL  = '2E75B6';
const LIGHT_FILL    = 'D6E4F0';
const GRID_FILL     = 'F2F7FB';
const WHITE         = 'FFFFFF';

const FONT = 'Arial';
const SZ   = 20;   // 10pt in half-points
const SZ_SM = 18;  // 9pt

const PAGE_W  = 15840;
const PAGE_H  = 12240;
const MARGIN  = 720;
const CONTENT_W = PAGE_W - MARGIN * 2; // 14400 DXA

const COL_LABEL = 2200;
const COL_DATA  = CONTENT_W - COL_LABEL; // 12200

// Skills columns
const COL_CAT  = 2880;
const COL_TITL = 4320;
const COL_STMT = 7200;

// CCT columns
const COL_CCT_THEME = 3000;
const COL_CCT_DESC  = 8200;
const COL_CCT_ALIGN = CONTENT_W - COL_CCT_THEME - COL_CCT_DESC; // 3200

// PO columns
const COL_PO_NAME  = 4000;
const COL_PO_ALIGN = CONTENT_W - COL_PO_NAME; // 10400

// ─── SLATE → DOCX HELPERS ─────────────────────────────────────────────────────

function slateToLines(slateContent) {
  if (!slateContent || typeof slateContent !== 'object') {
    const t = String(slateContent || '').replace(/\u00A0/g, ' ');
    return [{ style: 'Paragraph', segments: [{ text: t, bold: false, url: null }] }];
  }
  const root = slateContent.root || slateContent;
  const children = root.children || [];
  const rawLines = children.map(line => {
    const style = line.style || 'Paragraph';
    const segs = (line.children || []).map(child => {
      if (child.type === 'InlineStructuredValue') {
        const v = child.value || {};
        const url = v.url || null;
        const name = v.name || v.url || '';
        return { text: name, bold: false, italic: false, underline: false, url };
      }
      const txt = (child.text || '').replace(/\u00A0/g, ' ');
      return {
        text: txt,
        bold: !!child.bold,
        italic: !!child.italic,
        underline: !!child.underline,
        url: null
      };
    }).filter(s => s.text.trim().length > 0 || s.url);
    return { style, segments: segs };
  }).filter(line => line.segments.length > 0);

  // Expand any lines that have embedded \n into multiple lines
  const expanded = [];
  for (const line of rawLines) {
    const sub = expandEmbeddedNewlines(line);
    expanded.push(...sub);
  }
  return expanded;
}

function slateLinesToParagraphs(lines) {
  const paras = [];
  // Track whether we are "inside a header block" (= last non-empty line was a header)
  let lastNonEmptyWasHeader = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isBullet = line.style === 'BulletedList';
    const allText = line.segments.map(s => s.text).join('').trim();
    const isHeader = isSectionHeader(line.style, allText);

    // Skip lines that are purely whitespace/empty
    if (!allText && !line.segments.some(s => s.url)) continue;

    // ── SECTION HEADER ────────────────────────────────────────────────────────
    if (isHeader) {
      // Blank line BEFORE header (unless very first paragraph)
      if (paras.length > 0) {
        paras.push(makeBlankPara());
      }

      // Build header paragraph — ALL text forced bold, underline preserved, hyperlinks respected
      const children = [];
      for (const seg of line.segments) {
        if (seg.url) {
          children.push(new ExternalHyperlink({
            link: seg.url,
            children: [new TextRun({ text: seg.text, style: 'Hyperlink', font: FONT, size: SZ, bold: true })]
          }));
        } else {
          const opts = { font: FONT, size: SZ, bold: true };
          if (seg.underline) opts.underline = {};
          children.push(new TextRun({ text: seg.text, ...opts }));
        }
      }
      paras.push(new Paragraph({ children, spacing: { before: 0, after: 0 } }));

      // Blank line AFTER header
      paras.push(makeBlankPara());
      lastNonEmptyWasHeader = true;
      continue;
    }

    // ── BULLET OR PROMOTED-BULLET ─────────────────────────────────────────────
    // Paragraph-style lines get promoted to bullets if the preceding non-empty line was a header
    const promoteToBullet = isBullet || lastNonEmptyWasHeader;

    const children = [];
    for (const seg of line.segments) {
      if (seg.url) {
        children.push(new ExternalHyperlink({
          link: seg.url,
          children: [new TextRun({ text: seg.text, style: 'Hyperlink', font: FONT, size: SZ })]
        }));
        continue;
      }
      // Only apply bold/italic/underline if Coda explicitly set it
      const opts = { font: FONT, size: SZ };
      if (seg.bold)      opts.bold = true;
      if (seg.italic)    opts.italics = true;
      if (seg.underline) opts.underline = {};
      children.push(new TextRun({ text: seg.text, ...opts }));
    }

    const paraOpts = { children, spacing: { before: 0, after: 40 } };
    if (promoteToBullet) {
      paraOpts.bullet = { level: 0 };
    }
    paras.push(new Paragraph(paraOpts));
    lastNonEmptyWasHeader = false;
  }

  return paras.length ? paras : [new Paragraph({ children: [new TextRun({ text: '' })] })];
}

// A line is a section header if:
//  - it is NOT a BulletedList
//  - its FIRST segment's full text ends with ':' and is short, OR
//  - it matches a known standalone header pattern (with or without colon)
function isSectionHeader(style, allText) {
  if (style === 'BulletedList') return false;
  const t = allText.trim();
  if (!t) return false;
  // Ends with colon = label header
  if (t.endsWith(':') && t.length < 140) return true;
  // Known standalone headers without trailing colon
  if (/^(The learning experience should include the following|Out of Scope|COURSE LEVEL SCOPE NOTES|Course Notes|SOAP Scope Notes|Alignment with AACN Domains|Domain \d|Crosscutting Themes)/i.test(t)) return true;
  return false;
}

function makeBlankPara() {
  return new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: SZ })], spacing: { before: 0, after: 0 } });
}

// Expand a single slate line that contains embedded \n into multiple sub-lines,
// each inheriting the same style and segments (split proportionally)
function expandEmbeddedNewlines(line) {
  // Collect all text + segments, then split on \n
  // We need to split the segments array on \n boundaries
  const result = [];
  let currentSegs = [];

  for (const seg of line.segments) {
    if (!seg.url && seg.text.includes('\n')) {
      const parts = seg.text.split('\n');
      for (let pi = 0; pi < parts.length; pi++) {
        const part = parts[pi];
        if (pi === 0) {
          // First part goes on current line
          if (part.trim()) currentSegs.push({ ...seg, text: part });
        } else {
          // Flush current line
          if (currentSegs.length > 0) {
            result.push({ style: line.style, segments: currentSegs });
          }
          currentSegs = [];
          if (part.trim()) currentSegs.push({ ...seg, text: part });
        }
      }
    } else {
      currentSegs.push(seg);
    }
  }
  // Flush last
  if (currentSegs.length > 0) {
    result.push({ style: line.style, segments: currentSegs });
  }
  return result.length > 0 ? result : [line];
}

function parseInline(text, baseOpts = {}) {
  const children = [];
  const re = /\[([^\]]+)\]\((https?:\/\/[^)]+)\)|(https?:\/\/\S+)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) {
      const chunk = text.slice(last, m.index);
      if (chunk) children.push(new TextRun({ text: chunk, font: FONT, size: SZ, ...baseOpts }));
    }
    if (m[1]) {
      children.push(new ExternalHyperlink({ link: m[2], children: [new TextRun({ text: m[1], style: 'Hyperlink', font: FONT, size: SZ })] }));
    } else {
      children.push(new ExternalHyperlink({ link: m[3], children: [new TextRun({ text: m[3], style: 'Hyperlink', font: FONT, size: SZ })] }));
    }
    last = m.index + m[0].length;
  }
  if (last < text.length) {
    const rest = text.slice(last);
    if (rest) children.push(new TextRun({ text: rest, font: FONT, size: SZ, ...baseOpts }));
  }
  return children.length ? children : [new TextRun({ text, font: FONT, size: SZ, ...baseOpts })];
}


function slateLinesToParagraphs(lines) {
  const paras = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isBullet = line.style === 'BulletedList';
    const allText = line.segments.map(s => s.text).join('').trim();
    const isHeader = isSectionHeader(line.style, allText);

    // Skip lines that are purely whitespace/empty
    if (!allText && !line.segments.some(s => s.url)) continue;

    // ── SECTION HEADER ────────────────────────────────────────────────────────
    if (isHeader) {
      // Blank line BEFORE header (unless very first paragraph)
      if (paras.length > 0) {
        paras.push(makeBlankPara());
      }

      // Build header paragraph — ALL text forced bold, respect hyperlinks
      const children = [];
      for (const seg of line.segments) {
        if (seg.url) {
          children.push(new ExternalHyperlink({
            link: seg.url,
            children: [new TextRun({ text: seg.text, style: 'Hyperlink', font: FONT, size: SZ, bold: true })]
          }));
        } else {
          children.push(new TextRun({ text: seg.text, font: FONT, size: SZ, bold: true }));
        }
      }
      paras.push(new Paragraph({ children, spacing: { before: 0, after: 0 } }));

      // Blank line AFTER header
      paras.push(makeBlankPara());
      continue;
    }

    // ── BULLET OR PROMOTED-BULLET ─────────────────────────────────────────────
    // Promote Paragraph-style lines to bullets IF they follow a header
    // (look back past any blank paras to find whether the last real line was a header)
    let promoteToBullet = isBullet;
    if (!isBullet) {
      // Check if preceding non-empty line was a header
      for (let j = i - 1; j >= 0; j--) {
        const prev = lines[j];
        const prevText = prev.segments.map(s => s.text).join('').trim();
        if (!prevText) continue; // skip blanks
        if (isSectionHeader(prev.style, prevText)) {
          promoteToBullet = true;
        }
        break;
      }
    }

    const children = [];
    for (const seg of line.segments) {
      if (seg.url) {
        children.push(new ExternalHyperlink({
          link: seg.url,
          children: [new TextRun({ text: seg.text, style: 'Hyperlink', font: FONT, size: SZ })]
        }));
        continue;
      }

      // Only apply bold/italic/underline if Coda explicitly set it — never auto-bold
      const opts = { font: FONT, size: SZ };
      if (seg.bold)      opts.bold = true;
      if (seg.italic)    opts.italics = true;
      if (seg.underline) opts.underline = {};

      children.push(new TextRun({ text: seg.text, ...opts }));
    }

    const paraOpts = { children, spacing: { before: 0, after: 40 } };
    if (promoteToBullet) {
      paraOpts.bullet = { level: 0 };
    }
    paras.push(new Paragraph(paraOpts));
  }

  return paras.length ? paras : [new Paragraph({ children: [new TextRun({ text: '' })] })];
}

function plainTextToParagraphs(text) {
  if (!text || !text.trim()) return [new Paragraph({ children: [new TextRun({ text: '' })] })];
  return text.split('\n').map(line => {
    const t = line.replace(/\u00A0/g, ' ');
    const runs = parseInline(t);
    return new Paragraph({ children: runs, spacing: { after: 40 } });
  });
}

// ─── TABLE FACTORY ───────────────────────────────────────────────────────────────
function makeTable(rows, colWidths) {
  return new Table({
    rows,
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths
  });
}

// ─── TABLE CELL BUILDERS ──────────────────────────────────────────────────────

function headerCell(text, w, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: opts.fill || HEADER_FILL, type: ShadingType.SOLID },
    children: [new Paragraph({
      children: [new TextRun({ text, font: FONT, size: opts.sz || SZ, bold: true, color: WHITE })],
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 40, after: 40 }
    })],
    margins: { top: 60, bottom: 60, left: 80, right: 80 }
  });
}

function subHeaderCell(text, w, span = 1) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: SUBHEAD_FILL, type: ShadingType.SOLID },
    columnSpan: span,
    children: [new Paragraph({
      children: [new TextRun({ text, font: FONT, size: SZ, bold: true, color: WHITE })],
      spacing: { before: 40, after: 40 }
    })],
    margins: { top: 60, bottom: 60, left: 80, right: 80 }
  });
}

function lightCell(text, w, opts = {}) {
  // Label cells are always bold (matching the reference D948 format)
  const bold = opts.bold !== false;
  const runs = typeof text === 'string'
    ? parseInline(text, { bold })
    : [new TextRun({ text: String(text), font: FONT, size: SZ, bold })];
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: LIGHT_FILL, type: ShadingType.SOLID },
    columnSpan: opts.span || 1,
    children: [new Paragraph({
      children: runs,
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 40, after: 40 }
    })],
    margins: { top: 60, bottom: 60, left: 80, right: 80 }
  });
}

function dataCell(paragraphs, w, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.SOLID } : undefined,
    columnSpan: opts.span || 1,
    children: paragraphs,
    margins: { top: 60, bottom: 60, left: 80, right: 80 }
  });
}

function slateCellFull(slateOrText, w, span = 1) {
  let paras;
  if (typeof slateOrText === 'object' && slateOrText !== null && (slateOrText.root || slateOrText.type === 'slate')) {
    paras = slateLinesToParagraphs(slateToLines(slateOrText));
  } else if (typeof slateOrText === 'string') {
    paras = plainTextToParagraphs(slateOrText);
  } else {
    paras = [new Paragraph({ children: [new TextRun({ text: '' })] })];
  }
  return dataCell(paras, w, { span });
}

// ─── SPACER ───────────────────────────────────────────────────────────────────
function spacer() {
  return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 80 } });
}

// ─── SECTION LABEL ROW ────────────────────────────────────────────────────────
function bannerRow(text, fill, colSpan) {
  return new TableRow({
    children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { fill, type: ShadingType.SOLID },
      columnSpan: colSpan,
      children: [new Paragraph({
        children: [new TextRun({ text, font: FONT, size: SZ, bold: true, color: fill === LIGHT_FILL ? '000000' : WHITE })],
        spacing: { before: 40, after: 40 }
      })],
      margins: { top: 60, bottom: 60, left: 80, right: 80 }
    })]
  });
}

// ─── COURSE-LEVEL TABLES ──────────────────────────────────────────────────────

function buildScopeNotesTable() {
  const rows = [
    // Row 1: Dark-blue banner — two cells matching the data row's column widths
    new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell('Course Level Scope Notes', COL_LABEL),
      headerCell('', COL_DATA)
    ]}),
    // Row 2: Label + content
    new TableRow({ children: [
      lightCell('Scope Notes', COL_LABEL),
      slateCellFull(course.scopeNotes, COL_DATA)
    ]})
  ];
  return makeTable(rows, [COL_LABEL, COL_DATA]);
}

function buildAssessmentStrategyTable() {
  const modalityText = course.assessmentModality || '';
  const rationaleText = course.assessmentModalityRationale || '';
  const evidenceSlate = course.evidence;

  const rows = [
    new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell('Course Level Assessment Strategy', COL_LABEL),
      headerCell('Detail', COL_DATA)
    ]}),
    new TableRow({ children: [
      lightCell('Assessment Modality', COL_LABEL),
      dataCell([new Paragraph({ children: [new TextRun({ text: modalityText, font: FONT, size: SZ })], spacing: { after: 40 } })], COL_DATA)
    ]}),
    new TableRow({ children: [
      lightCell('Modality Rationale', COL_LABEL),
      dataCell(plainTextToParagraphs(rationaleText), COL_DATA)
    ]}),
    new TableRow({ children: [
      lightCell('Evidence', COL_LABEL),
      slateCellFull(evidenceSlate, COL_DATA)
    ]})
  ];
  if (course.asmtNotes && course.asmtNotes.trim()) {
    rows.push(new TableRow({ children: [
      lightCell('Assessment Notes', COL_LABEL),
      dataCell(plainTextToParagraphs(course.asmtNotes), COL_DATA)
    ]}));
  }
  return makeTable(rows, [COL_LABEL, COL_DATA]);
}

function buildLRStrategyTable() {
  return makeTable([
      new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell('Course Level LR Strategy', COL_LABEL), headerCell('Detail', COL_DATA)] }),
      new TableRow({ children: [lightCell('LR Strategy', COL_LABEL), slateCellFull(course.lrStrategy, COL_DATA)] })
    ], [COL_LABEL, COL_DATA]);
}

function buildToolsTable() {
  return makeTable([
      new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell('Course Level Critical Tools and Technologies', COL_LABEL), headerCell('Detail', COL_DATA)] }),
      new TableRow({ children: [lightCell('Tools & Technologies', COL_LABEL), slateCellFull(course.tools, COL_DATA)] })
    ], [COL_LABEL, COL_DATA]);
}

// ─── CROSS CUTTING THEMES TABLE ───────────────────────────────────────────────

function buildCCTTable(ccts) {
  if (!ccts || ccts.length === 0) return null;

  const rows = [
    // Row 1: Dark-blue banner — three cells matching column widths
    new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell('Course Level Cross Cutting Themes', COL_CCT_THEME),
      headerCell('', COL_CCT_DESC),
      headerCell('', COL_CCT_ALIGN)
    ]}),
    // Row 2: Medium-blue column headers
    new TableRow({ children: [
      subHeaderCell('Cross Cutting Theme', COL_CCT_THEME),
      subHeaderCell('Description', COL_CCT_DESC),
      subHeaderCell('Aligned', COL_CCT_ALIGN)
    ]})
  ];

  for (const cct of ccts) {
    rows.push(new TableRow({ children: [
      dataCell([new Paragraph({ children: [new TextRun({ text: cct.name || '', font: FONT, size: SZ })], spacing: { after: 40 } })], COL_CCT_THEME),
      dataCell([new Paragraph({ children: [new TextRun({ text: cct.description || '', font: FONT, size: SZ })], spacing: { after: 40 } })], COL_CCT_DESC),
      dataCell([new Paragraph({ children: [new TextRun({ text: cct.aligned || '', font: FONT, size: SZ, bold: true })], spacing: { after: 40 } })], COL_CCT_ALIGN)
    ]}));
  }

  return makeTable(rows, [COL_CCT_THEME, COL_CCT_DESC, COL_CCT_ALIGN]);
}

function buildPOTable(pos) {
  if (!pos || pos.length === 0) return null;

  const rows = [
    // Row 1: Dark-blue banner — two cells matching column widths
    new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell('Course Level Program Outcomes', COL_PO_NAME),
      headerCell('', COL_PO_ALIGN)
    ]}),
    // Row 2: Medium-blue column headers
    new TableRow({ children: [
      subHeaderCell('Program Outcome Display', COL_PO_NAME),
      subHeaderCell('Aligned', COL_PO_ALIGN)
    ]})
  ];

  for (const po of pos) {
    rows.push(new TableRow({ children: [
      dataCell([new Paragraph({ children: [new TextRun({ text: po.name || '', font: FONT, size: SZ })], spacing: { after: 40 } })], COL_PO_NAME),
      dataCell([new Paragraph({ children: [new TextRun({ text: po.aligned || '', font: FONT, size: SZ, bold: true })], spacing: { after: 40 } })], COL_PO_ALIGN)
    ]}));
  }

  return makeTable(rows, [COL_PO_NAME, COL_PO_ALIGN]);
}

function buildCompTable(comp, idx) {
  const titleParts = splitCompTitle(comp.titleAndStatement);
  const title = titleParts.title;
  const statement = titleParts.statement;
  const levelText = comp.level ? `Target Level: ${comp.level}` : '';

  // Row 1: Blue header
  const headerText = [
    `Competency ${idx}`,
    title ? `| Title: ${title}` : '',
    statement ? `| Statement: ${statement}` : '',
    levelText ? `| ${levelText}` : ''
  ].filter(Boolean).join(' ');

  const rows = [
    new TableRow({ tableHeader: true, height: { value: 400, rule: HeightRule.ATLEAST }, children: [
      headerCell(headerText, CONTENT_W, { span: 3 })] }),
    bannerRow('Skills', LIGHT_FILL, 3),
    new TableRow({ children: [
      subHeaderCell('Skill Category', COL_CAT),
      subHeaderCell('Skill Title', COL_TITL),
      subHeaderCell('Skill Statement', COL_STMT)
    ]})
  ];

  if (comp.skills && comp.skills.length > 0) {
    for (const skill of comp.skills) {
      rows.push(new TableRow({ children: [
        dataCell([new Paragraph({ children: [new TextRun({ text: skill.category || '', font: FONT, size: SZ })], spacing: { after: 40 } })], COL_CAT, { shading: GRID_FILL }),
        dataCell([new Paragraph({ children: [new TextRun({ text: skill.title || '', font: FONT, size: SZ })], spacing: { after: 40 } })], COL_TITL),
        dataCell([new Paragraph({ children: [new TextRun({ text: skill.statement || '', font: FONT, size: SZ })], spacing: { after: 40 } })], COL_STMT)
      ]}));
    }
  } else {
    rows.push(new TableRow({ children: [
      dataCell([new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: SZ })] })], CONTENT_W, { span: 3 })
    ]}));
  }

  // Competency Detail section
  rows.push(bannerRow('Competency Detail', LIGHT_FILL, 3));

  // Assessment Recommendations
  const modalityText = comp.modality || '';
  const rationaleText = comp.modalityRationale || '';
  const evidenceText = comp.evidence || '';
  const asmtContent = [
    modalityText && `Modality: ${modalityText}`,
    rationaleText && `Rationale: ${rationaleText}`,
    evidenceText && `Evidence: ${evidenceText}`
  ].filter(Boolean).join('\n');

  rows.push(new TableRow({ children: [
    lightCell('Assessment Recommendations', COL_LABEL),
    dataCell(plainTextToParagraphs(asmtContent), COL_DATA, { span: 2 })
  ]}));

  // Additional Scope Notes
  const scopeText = comp.scopeNotes || '';
  rows.push(new TableRow({ children: [
    lightCell('Additional Scope Notes and Requirements', COL_LABEL),
    dataCell(plainTextToParagraphs(scopeText), COL_DATA, { span: 2 })
  ]}));

  // Standards Alignment
  const stdSlate = comp.standardsAlignment;
  rows.push(new TableRow({ children: [
    lightCell('Standards Alignment', COL_LABEL),
    slateCellFull(stdSlate, COL_DATA, 2)
  ]}));

  return makeTable(rows, [COL_CAT, COL_TITL, COL_STMT]);
}

function splitCompTitle(raw) {
  if (!raw) return { title: '', statement: '' };
  // Try ".  The learner" or ". The learner"
  const patterns = ['.  The learner', '. The learner', '.  The student', '. The student'];
  for (const p of patterns) {
    const idx = raw.indexOf(p);
    if (idx !== -1) {
      return {
        title: raw.slice(0, idx).trim(),
        statement: raw.slice(idx + 2).trim()
      };
    }
  }
  return { title: raw.trim(), statement: '' };
}

// ─── DOCUMENT ASSEMBLY ────────────────────────────────────────────────────────

function buildDocument() {
  const courseCode = (course.code || '').trim();
  const courseName = (course.name || '').trim();
  const modality   = course.modality || '';
  const status     = course.status || '';
  const units      = course.units != null ? String(course.units) : '';

  // QA log
  console.log('=== QA CHECK ===');
  console.log(`Course: ${courseCode} – ${courseName}`);
  console.log(`Modality: ${modality} | Status: ${status} | Units: ${units}`);
  console.log(`Scope Notes chars: ${JSON.stringify(course.scopeNotes).length}`);
  console.log(`Assessment Modality: ${course.assessmentModality}`);
  console.log(`LR Strategy chars: ${JSON.stringify(course.lrStrategy).length}`);
  console.log(`Tools chars: ${JSON.stringify(course.tools).length}`);
  console.log(`Competencies: ${comps.length}`);
  console.log(`CCTs (non-blank aligned): ${ccts.length}`);
  console.log(`Program Outcomes (non-blank aligned): ${pos.length}`);
  if (ccts.length) ccts.forEach(c => console.log(`  CCT: ${c.name} → ${c.aligned}`));
  if (pos.length) pos.forEach(p => console.log(`  PO: ${p.name} → ${p.aligned}`));
  console.log('=== END QA ===');

  const titleLine = `${courseCode}${courseCode ? ' – ' : ''}${courseName}`;
  const subtitleLine = 'Competency Context Worksheet';
  const infoLine = [
    modality && `Modality: ${modality}`,
    status && `Status: ${status}`,
    units && `Credit Units: ${units}`
  ].filter(Boolean).join('  |  ');

  // Page header
  const pageHeader = new Header({
    children: [new Paragraph({
      children: [
        new TextRun({ text: `${courseCode} – ${courseName}  |  Page `, font: FONT, size: SZ_SM }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ_SM })
      ],
      alignment: AlignmentType.RIGHT
    })]
  });

  const children = [];

  // ── Title block (CENTERED) ──────────────────────────────────────────────────
  children.push(new Paragraph({
    children: [new TextRun({ text: titleLine, font: FONT, size: 28, bold: true })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 }
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: subtitleLine, font: FONT, size: 24, bold: false })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 }
  }));
  if (infoLine) {
    children.push(new Paragraph({
      children: [new TextRun({ text: infoLine, font: FONT, size: SZ })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 }
    }));
  }

  // ── Course-level tables ─────────────────────────────────────────────────────
  children.push(buildScopeNotesTable());
  children.push(spacer());

  children.push(buildAssessmentStrategyTable());
  children.push(spacer());

  children.push(buildLRStrategyTable());
  children.push(spacer());

  children.push(buildToolsTable());
  children.push(spacer());

  // ── Cross Cutting Themes (NEW) ──────────────────────────────────────────────
  const cctTable = buildCCTTable(ccts);
  if (cctTable) {
    children.push(cctTable);
    children.push(spacer());
  }

  // ── Program Outcomes (NEW) ──────────────────────────────────────────────────
  const poTable = buildPOTable(pos);
  if (poTable) {
    children.push(poTable);
    children.push(spacer());
  }

  // ── Competency section heading ───────────────────────────────────────────────
  children.push(new Paragraph({
    children: [new TextRun({ text: 'Course Competencies and Context Tables', font: FONT, size: 22, bold: true })],
    spacing: { before: 120, after: 120 }
  }));

  // ── Competency tables ───────────────────────────────────────────────────────
  for (let i = 0; i < comps.length; i++) {
    children.push(buildCompTable(comps[i], i + 1));
    children.push(spacer());
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: SZ }
        }
      },
      characterStyles: [{
        id: 'Hyperlink',
        name: 'Hyperlink',
        run: { color: '0563C1', underline: { type: 'single' } }
      }]
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.LANDSCAPE },
          margin: { top: 1080, right: MARGIN, bottom: MARGIN, left: MARGIN, header: 360 }
        }
      },
      headers: { default: pageHeader },
      children
    }]
  });

  return doc;
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
(async () => {
  try {
    const doc = buildDocument();
    const courseCode = (DATA.course.code || 'COURSE').replace(/\s+/g, '').toUpperCase();
    const outPath = path.join('/mnt/user-data/outputs', `${courseCode}_CCW.docx`);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buffer);
    console.log(`Written: ${outPath} (${(buffer.length / 1024).toFixed(1)} KB)`);
  } catch (err) {
    console.error('BUILD FAILED:', err);
    process.exit(1);
  }
})();
