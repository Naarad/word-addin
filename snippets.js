/*
 * Thoughtrons Word add-in — snippet definitions.
 *
 * Each snippet is a function that uses the Word API to insert a styled
 * block at the cursor (or at the end of the document if no selection).
 *
 * Adding a new snippet: append to SNIPPETS, then add it to the taskpane's
 * LIBRARY array (src/taskpane.html). New snippets show up in the library
 * automatically — no manifest change needed.
 *
 * Brand tokens — keep in sync with the .docx templates and PowerPoint add-in.
 */

window.ThoughtronsBrand = {
  ORANGE: "#ED8B1F",
  TEXT:   "#111111",
  TEXT2:  "#555555",
  TEXT3:  "#888888",
  RULE:   "#DDDDDD",
  HEADER_FILL: "#1F1F1F",
};

const C = window.ThoughtronsBrand;

// ----- Helpers -----
function ipara(body, text, opts = {}) {
  const p = body.insertParagraph(text, "End");
  p.font.name = opts.font || "Calibri";
  p.font.size = opts.size || 11;
  p.font.bold = !!opts.bold;
  p.font.italic = !!opts.italic;
  p.font.color = opts.color || C.TEXT;
  if (opts.spacingBefore != null) p.spaceBefore = opts.spacingBefore;
  if (opts.spacingAfter != null)  p.spaceAfter  = opts.spacingAfter;
  if (opts.style) p.styleBuiltIn = opts.style;
  if (opts.align) p.alignment = opts.align;
  return p;
}

function ieyebrow(body, text) {
  return ipara(body, text.toUpperCase(),
    { font: "Consolas", size: 9, bold: true, color: C.ORANGE,
      spacingBefore: 8, spacingAfter: 4 });
}

function ititle(body, text) {
  return ipara(body, text,
    { font: "Calibri", size: 28, bold: true, color: C.TEXT,
      spacingBefore: 0, spacingAfter: 8 });
}

function ih1(body, text) {
  return ipara(body, text,
    { font: "Calibri", size: 16, bold: true, color: C.TEXT,
      spacingBefore: 16, spacingAfter: 6, style: "Heading1" });
}

function ih2(body, text) {
  return ipara(body, text,
    { font: "Calibri", size: 13, bold: true, color: C.TEXT,
      spacingBefore: 10, spacingAfter: 4, style: "Heading2" });
}

function ibody(body, text) {
  return ipara(body, text,
    { font: "Calibri", size: 11, color: C.TEXT,
      spacingAfter: 4 });
}

// Insert a styled table. data[0] is the header row.
async function itable(context, data, options = {}) {
  const body = context.document.body;
  const tbl = body.insertTable(data.length, data[0].length, "End", data);
  tbl.styleBuiltIn = "PlainTable1";
  context.load(tbl, "rows/items/cells/items/body/font");
  await context.sync();

  // Header row — dark fill, white mono text
  const headerCells = tbl.rows.items[0].cells;
  headerCells.forEach(cell => {
    cell.shadingColor = C.HEADER_FILL;
    cell.body.font.color = "#FFFFFF";
    cell.body.font.name = "Consolas";
    cell.body.font.size = 9;
    cell.body.font.bold = true;
  });

  // Body rows — Calibri, dark text
  for (let r = 1; r < tbl.rows.items.length; r++) {
    const cells = tbl.rows.items[r].cells;
    cells.forEach(cell => {
      cell.body.font.color = C.TEXT;
      cell.body.font.name = "Calibri";
      cell.body.font.size = 11;
    });
  }
  return tbl;
}

// ============================================================
// SNIPPETS — append new ones here and register in taskpane.html
// ============================================================
window.ThoughtronsSnippets = {

  // ---------- Headings & marks ----------
  eyebrow: async (ctx) => {
    ieyebrow(ctx.document.body, "// Section  /  Topic");
    return ctx.sync();
  },

  title: async (ctx) => {
    ititle(ctx.document.body, "Document title goes here.");
    return ctx.sync();
  },

  h1: async (ctx) => {
    ih1(ctx.document.body, "Section heading");
    return ctx.sync();
  },

  h2: async (ctx) => {
    ih2(ctx.document.body, "Sub-section heading");
    return ctx.sync();
  },

  // ---------- Common content blocks ----------
  exec_summary: async (ctx) => {
    const body = ctx.document.body;
    ih1(body, "Executive Summary");
    ibody(body, "Two to four sentences. Lead with the takeaway. Mention the headline number. Reference the trust signal. Close with the timeline.");
    return ctx.sync();
  },

  context_block: async (ctx) => {
    const body = ctx.document.body;
    ih1(body, "Context");
    ibody(body, "State the situation in one paragraph. What we're trying to achieve, what we're up against, what's preventing us from doing it ourselves.");
    return ctx.sync();
  },

  callout: async (ctx) => {
    const body = ctx.document.body;
    const p = ipara(body, "Key takeaway in one short sentence.",
      { font: "Calibri", size: 13, italic: true, color: C.ORANGE,
        spacingBefore: 12, spacingAfter: 12, bold: true });
    return ctx.sync();
  },

  // ---------- Tables ----------
  action_items_table: async (ctx) => {
    await itable(ctx, [
      ["#",  "Action",                          "Owner",  "Due",     "Status"],
      ["A1", "Action — verb-first, specific.",  "Name",   "DD MMM",  "OPEN"],
      ["A2", "Action.",                          "Name",   "DD MMM",  "OPEN"],
      ["A3", "Action.",                          "Name",   "DD MMM",  "OPEN"],
    ]);
    return ctx.sync();
  },

  decisions_table: async (ctx) => {
    await itable(ctx, [
      ["#",  "Decision",                         "Owner",  "Effective"],
      ["D1", "Decision — concrete and specific.","Name",   "Immediate"],
      ["D2", "Decision.",                         "Name",   "DD MMM"],
    ]);
    return ctx.sync();
  },

  risk_table: async (ctx) => {
    await itable(ctx, [
      ["Risk",                       "Impact",  "Likelihood", "Owner",  "Mitigation"],
      ["Risk description one.",      "HIGH",    "MED",        "Name",   "Mitigation in 1-2 lines."],
      ["Risk description two.",      "MED",     "HIGH",       "Name",   "Mitigation in 1-2 lines."],
      ["Risk description three.",    "LOW",     "MED",        "Name",   "Mitigation in 1-2 lines."],
    ]);
    return ctx.sync();
  },

  pricing_table: async (ctx) => {
    await itable(ctx, [
      ["Item",                   "Fee (₹)"],
      ["Discovery",              "X,XX,XXX"],
      ["Design",                 "X,XX,XXX"],
      ["Plan + hand-off",        "X,XX,XXX"],
      ["Subtotal",               "X,XX,XXX"],
      ["Taxes (as applicable)",  "—"],
      ["Total",                  "X,XX,XXX"],
    ]);
    return ctx.sync();
  },

  // ---------- Reference / academic blocks ----------
  references_block: async (ctx) => {
    const body = ctx.document.body;
    const p = ipara(body, "REFERENCES",
      { font: "Times New Roman", size: 11, bold: true, color: C.TEXT,
        spacingBefore: 16, spacingAfter: 8, align: "Centered" });
    [1,2,3].forEach(n => {
      ipara(body, `[${n}] Author, A., "Title of cited work," Venue, year.`,
        { font: "Times New Roman", size: 9, color: C.TEXT, spacingAfter: 2 });
    });
    return ctx.sync();
  },

  figure_caption: async (ctx) => {
    const body = ctx.document.body;
    const p = body.insertParagraph("Fig. N. ", "End");
    p.font.name = "Times New Roman";
    p.font.size = 9;
    p.font.italic = true;
    p.font.bold = true;
    p.font.color = C.TEXT;
    p.spaceBefore = 2;
    p.spaceAfter = 8;
    return ctx.sync();
  },

  // ---------- Memo header ----------
  memo_header: async (ctx) => {
    await itable(ctx, [
      ["TO",      "Recipients"],
      ["FROM",    "Author Name, Role"],
      ["DATE",    "DD MMM YYYY"],
      ["SUBJECT", "Memo subject"],
    ]);
    return ctx.sync();
  },

  // ---------- Classification stamps ----------
  classify_unclassified: (ctx) => stampClassification(ctx, "UNCLASSIFIED",  "#DCF5DC", "#2D7A2D"),
  classify_internal:     (ctx) => stampClassification(ctx, "INTERNAL",      "#FFE8C2", "#9B5A00"),
  classify_restricted:   (ctx) => stampClassification(ctx, "RESTRICTED",    "#FFD6D6", "#A12626"),
  classify_confidential: (ctx) => stampClassification(ctx, "CONFIDENTIAL",  "#FFD6D6", "#A12626"),
};

async function stampClassification(ctx, label, fill, fg) {
  // Insert into header AND footer of every section
  const sections = ctx.document.sections;
  ctx.load(sections, "items");
  await ctx.sync();
  for (const sec of sections.items) {
    const header = sec.getHeader("Primary");
    const footer = sec.getFooter("Primary");
    header.clear(); footer.clear();
    [header, footer].forEach(part => {
      const p = part.insertParagraph(label, "End");
      p.alignment = "Centered";
      p.font.name = "Consolas";
      p.font.size = 11;
      p.font.bold = true;
      p.font.color = fg;
    });
  }
  return ctx.sync();
}
