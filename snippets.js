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

  // ---------- Monthly updates (full-template inserts) ----------
  monthly_update_tech:                  (ctx) => insertMonthlyUpdate(ctx, MU_TECH),
  monthly_update_sales:                 (ctx) => insertMonthlyUpdate(ctx, MU_SALES),
  monthly_update_commercial_compliance: (ctx) => insertMonthlyUpdate(ctx, MU_COMM_COMPLIANCE),
  monthly_update_marketing:             (ctx) => insertMonthlyUpdate(ctx, MU_MARKETING),
  monthly_update_bd:                    (ctx) => insertMonthlyUpdate(ctx, MU_BD),

  // ---------- Full-document templates (inserted via base64 .docx) ----------
  research_paper:        (ctx) => insertTemplate(ctx, "research_paper"),
  proposal:              (ctx) => insertTemplate(ctx, "proposal"),
  programme_brief:       (ctx) => insertTemplate(ctx, "programme_brief"),
  internal_memo:         (ctx) => insertTemplate(ctx, "internal_memo"),
  meeting_minutes:       (ctx) => insertTemplate(ctx, "meeting_minutes"),
  monthly_report:        (ctx) => insertTemplate(ctx, "monthly_report"),
  win_loss_analysis:     (ctx) => insertTemplate(ctx, "win_loss_analysis"),
  press_release:         (ctx) => insertTemplate(ctx, "press_release"),
  case_study:            (ctx) => insertTemplate(ctx, "case_study"),
  blog_post:             (ctx) => insertTemplate(ctx, "blog_post"),
  linkedin_post:         (ctx) => insertTemplate(ctx, "linkedin_post"),
  campaign_plan:         (ctx) => insertTemplate(ctx, "campaign_plan"),
};

// ============================================================
// FULL-TEMPLATE INSERTER — uses Word's insertFileFromBase64 API
// ============================================================
// The base64 strings live in templates.js (loaded alongside snippets.js).
// To add a new template: add the .docx to Brand_System_v1/, add a (key, path)
// row to encode_templates.py, run it, then register the key here + in
// taskpane.html's LIBRARY array.
function insertTemplate(ctx, key) {
  const b64 = (window.ThoughtronsTemplates || {})[key];
  if (!b64) {
    return Promise.reject(new Error("Unknown template: " + key +
      " — make sure templates.js is loaded and the key exists."));
  }
  ctx.document.body.insertFileFromBase64(b64, "End");
  return ctx.sync();
}

// ============================================================
// MONTHLY UPDATE BUILDER (shared by all 5 department configs)
// ============================================================
async function insertMonthlyUpdate(ctx, cfg) {
  const body = ctx.document.body;

  // Header
  ieyebrow(body, `// Monthly Update / ${cfg.dept} / MONTH YYYY`);
  ititle(body, `Monthly Update — ${cfg.deptName}.`);
  ipara(body, "Author Name  ·  Role  ·  Period: MMM YYYY  ·  Filed: DD MMM YYYY",
    { font: "Consolas", size: 9, color: C.TEXT2, spacingAfter: 12 });

  // Meta block (4-cell table)
  await itable(ctx, [
    ["AUTHOR", "ROLE", "MANAGER", "PERIOD"],
    ["Author Name", "Role / Designation", "Manager Name", "MMM YYYY"],
  ]);

  // 1. HIGHLIGHTS
  ih1(body, "1. Highlights");
  ipara(body, "Top 3-5 things from this month. The reader should walk away knowing what mattered after this section alone.",
    { font: "Calibri", size: 11, italic: true, color: C.TEXT2, spacingAfter: 6 });
  cfg.exampleHighlights.forEach(h => ibullet(body, h));

  // 2. KEY NUMBERS
  ih1(body, "2. Key Numbers");
  ipara(body, `The four numbers that show your ${cfg.deptName.toLowerCase()} month at a glance.`,
    { font: "Calibri", size: 11, italic: true, color: C.TEXT2, spacingAfter: 6 });
  await itable(ctx, [
    ["Metric", "This month", "Last month", "Δ", "Notes"],
    [cfg.kpiLabels[0], "value", "value", "+/-X%", "One-line context"],
    [cfg.kpiLabels[1], "value", "value", "+/-X%", "One-line context"],
    [cfg.kpiLabels[2], "value", "value", "+/-X%", "One-line context"],
    [cfg.kpiLabels[3], "value", "value", "+/-X%", "One-line context"],
  ]);

  // 3+ Department-specific sections
  cfg.sections.forEach((s, i) => {
    ih1(body, `${i + 3}. ${s.heading}`);
    if (s.note)  ipara(body, s.note, { font: "Calibri", size: 11, italic: true, color: C.TEXT2, spacingAfter: 6 });
    if (s.body)  ibody(body, s.body);
    if (s.bullets) s.bullets.forEach(b => ibullet(body, b));
  });

  // Wins & Lessons
  const n = cfg.sections.length;
  ih1(body, `${n + 3}. Wins & Lessons`);
  ih2(body, "What worked");
  ibullet(body, "Specific thing one — and what about it worked.");
  ibullet(body, "Specific thing two.");
  ih2(body, "What didn't");
  ibullet(body, "Specific thing one — and what you'd do differently.");
  ibullet(body, "Specific thing two.");

  // Blockers & Asks
  ih1(body, `${n + 4}. Blockers & Asks`);
  ipara(body, "What's preventing you from going faster — and what specifically you need from your manager / leadership / a partner team to unblock.",
    { font: "Calibri", size: 11, italic: true, color: C.TEXT2, spacingAfter: 6 });
  ibullet(body, "Blocker one — what you need (decision / resource / introduction / approval) — by when.");
  ibullet(body, "Blocker two.");

  // Next Month
  ih1(body, `${n + 5}. Next Month`);
  ipara(body, "Top 3-5 priorities for next month. Be specific enough that you can audit yourself against this list when you write next month's update.",
    { font: "Calibri", size: 11, italic: true, color: C.TEXT2, spacingAfter: 6 });
  ibullet(body, "Priority one — concrete, time-boxed.");
  ibullet(body, "Priority two.");
  ibullet(body, "Priority three.");

  // Notes
  ih1(body, `${n + 6}. Notes`);
  ipara(body, "Anything else worth knowing — career development, leave plans, team observations, ideas you'd like time to explore.",
    { font: "Calibri", size: 11, color: C.TEXT2 });

  return ctx.sync();
}

// Helper for bulleted list items
function ibullet(body, text) {
  const p = body.insertParagraph(text, "End");
  p.font.name = "Calibri";
  p.font.size = 11;
  p.font.color = C.TEXT;
  p.spaceAfter = 2;
  p.styleBuiltIn = "ListBullet";
  return p;
}

// ============================================================
// DEPARTMENT CONFIGS for the monthly-update snippet
// ============================================================
const MU_TECH = {
  dept: "TECH", deptName: "Tech",
  exampleHighlights: [
    "Shipped CHAITANYA optical-design v0.3 — 8% weight reduction vs. v0.2.",
    "Closed integration of LWIR module with the YOLOv8n detector — 17 ms inference on M3 Max.",
    "Spun up CI for the firmware repo — reduced manual test cycles by 3 hours/week.",
  ],
  kpiLabels: ["Issues closed", "PRs merged", "Tests written", "Open programmes"],
  sections: [
    { heading: "Programmes & deliverables",
      note: "What you shipped, on which programmes, with what evidence. One sub-section per programme you contributed to.",
      bullets: ["Programme A — what you delivered (design, prototype, test, doc) — link to evidence.",
                "Programme B — same.", "Programme C — same."] },
    { heading: "Engineering quality",
      note: "Anything that improved the quality of what we ship — tests, CI, monitoring, code review, documentation.",
      bullets: ["Tests added / coverage delta.", "CI / tooling improvements.", "Tech debt addressed."] },
    { heading: "Experiments & learning",
      note: "What you tried that didn't ship — including failures. R&D progress is uneven by definition; document the negative results.",
      body: "Free-form section. Describe what you investigated, what you learned, whether it warrants more time." },
    { heading: "Skills & development",
      note: "What you learned this month. Courses, papers, internal sessions, conferences, mentoring.",
      bullets: ["Course / paper / book — what you took from it.", "Mentor sessions — with whom, on what.",
                "Conferences / talks — attended or delivered."] },
  ],
};

const MU_SALES = {
  dept: "SALES", deptName: "Sales",
  exampleHighlights: [
    "Closed-Won DPSU-A advisory engagement — ₹3.4Cr, 6-month engagement.",
    "Pipeline grew +38% QoQ; weighted pipeline at ₹14.2Cr.",
    "Lost capability assessment with Client X on price — initiated learning doc.",
  ],
  kpiLabels: ["New leads", "Meetings held", "Proposals sent", "Closed-won (₹)"],
  sections: [
    { heading: "Pipeline movement",
      note: "Deals that progressed, stalled, or shifted stage this month. Not a list of every conversation — only the consequential ones.",
      bullets: ["Deal A — moved from Qualified → Proposal — value, expected close.",
                "Deal B — stalled at Negotiation — what's holding it up.",
                "Deal C — slipped expected close date — by how long, why."] },
    { heading: "Closed deals",
      note: "Every Closed-Won and Closed-Lost this month. For every Closed-Lost, the reason and the lesson.",
      bullets: ["WIN — Client / programme / value / why we won.",
                "LOSS — Client / programme / why — link to win/loss analysis doc."] },
    { heading: "New leads & sources",
      note: "Where this month's leads came from. Useful for reviewing channel ROI quarterly.",
      bullets: ["Source X — count of leads — qualified count.", "Source Y — count — qualified count."] },
    { heading: "Customer feedback themes",
      note: "Recurring objections, requests, or signals you heard from prospects this month. Three or four themes max.",
      body: "Theme 1: <observation, frequency, what it suggests we change>. Theme 2: <same>." },
    { heading: "Top accounts to watch",
      note: "Two or three accounts you'd flag for leadership attention next month.",
      bullets: ["Account A — why it matters this month.", "Account B — why."] },
  ],
};

const MU_COMM_COMPLIANCE = {
  dept: "COMMERCIAL & COMPLIANCE", deptName: "Commercial & Compliance",
  exampleHighlights: [
    "Closed FOSS MoU — partnership terms favourable to Thoughtrons IP.",
    "Passed annual data-protection audit — zero findings.",
    "Renewed three vendor contracts; exited one on poor performance.",
  ],
  kpiLabels: ["Contracts processed", "Compliance items closed", "Audits cleared", "Vendor reviews"],
  sections: [
    { heading: "Contracts & legal",
      note: "Every contract that started, was renewed, amended, or terminated this month.",
      bullets: ["New — counterparty / type / value / key term.", "Renewed — counterparty / changes.",
                "Terminated / exited — counterparty / reason."] },
    { heading: "Compliance status",
      note: "DPP / certifications / audits / regulatory requirements. RAG status against each.",
      bullets: ["DPP — current status — next gate.", "ISO / quality cert — current status — next audit.",
                "Data protection — items closed / open."] },
    { heading: "Vendor & supplier management",
      note: "Onboarded vendors, off-boarded vendors, performance reviews completed.",
      bullets: ["Vendor A — review outcome.", "Vendor B — onboarded / off-boarded."] },
    { heading: "Procurement & cost",
      note: "Major procurement decisions, cost variances vs. plan, savings realised.",
      body: "Free-form. Material POs, cost surprises, savings, budget pressure points." },
    { heading: "Regulatory & policy updates",
      note: "Changes in MoD / DPP / export-control / data policy that affect Thoughtrons. One paragraph per relevant change.",
      body: "Update 1: <what changed, our exposure, action we're taking>." },
    { heading: "Risk register changes",
      note: "Risks added, retired, or re-scored this month. Link to the Risk Register file.",
      bullets: ["Added — risk description / score.", "Retired — what changed.",
                "Re-scored — old score → new score, why."] },
  ],
};

const MU_MARKETING = {
  dept: "MARKETING", deptName: "Marketing",
  exampleHighlights: [
    "Published 3 long-form posts; LinkedIn impressions +52% MoM.",
    "Press pickup in Industry Today — 'Mil-tech 25-under-25' feature.",
    "DefExpo booth 80% planned — demo unit and collateral on track.",
  ],
  kpiLabels: ["Content published", "Reach (impressions)", "Engagement rate", "Press pickups"],
  sections: [
    { heading: "Content published",
      note: "By channel. Counts and link to each piece.",
      bullets: ["Website / blog — count — top performer + URL.", "LinkedIn — count — top performer.",
                "Other channels — count."] },
    { heading: "Campaigns",
      note: "Active or completed campaigns — goal, audience, outcome vs. target.",
      bullets: ["Campaign A — goal / audience / status / metric vs. target.", "Campaign B — same."] },
    { heading: "Press & media coverage",
      note: "Earned press, podcast appearances, third-party features. Inbound interview requests too.",
      bullets: ["Outlet — date — angle — reach estimate.",
                "Pending requests — what we said yes / no to and why."] },
    { heading: "Events & exhibitions",
      note: "Events worked on this month — pre-event prep, on-site activities, post-event follow-up.",
      bullets: ["Event A — phase — what was done — what's next."] },
    { heading: "Brand assets & website",
      note: "Updates to deck templates, brand guide, website. Anything the rest of the org now has access to.",
      body: "What was created / updated, and where to find it." },
    { heading: "Social engagement themes",
      note: "Recurring comments / DMs / requests on our channels this month. Useful for content team and product team.",
      body: "Theme 1: <observation>. Theme 2: <observation>." },
  ],
};

const MU_BD = {
  dept: "BD", deptName: "BD",
  exampleHighlights: [
    "Signed MoU with FOSS — joint go-to-market on thermal sensor integration.",
    "DPSU-A: champion identified; programme XYZ on radar for FY27 budget.",
    "Attended DefExpo briefing — 4 conversations qualified, 1 pushed to Sales as lead.",
  ],
  kpiLabels: ["Active partners", "MoUs signed", "Strategic accounts engaged", "Govt / DPSU touchpoints"],
  sections: [
    { heading: "Partnership pipeline",
      note: "Status by partner. Different from sales pipeline — these are joint-venture / MoU / co-development relationships.",
      bullets: ["Partner A — stage — what's pending — owner.", "Partner B — stage."] },
    { heading: "New partner conversations",
      note: "First-touch conversations with potential partners this month. Even if they go nowhere, log them.",
      bullets: ["Org X — context of intro — outcome — follow-up.", "Org Y — same."] },
    { heading: "Strategic accounts",
      note: "DPSUs, primes, large customers — engagement update per account. Link to Account Maps.",
      bullets: ["Account 1 — change in champion / blocker / programme this month.", "Account 2 — change."] },
    { heading: "Government & DPSU engagement",
      note: "Briefings, RFP awareness, programme intelligence, policy interactions.",
      bullets: ["Engagement A — context, outcome.", "Engagement B — same."] },
    { heading: "Conferences & industry events",
      note: "Events you attended or spoke at this month. What was learned, who was met.",
      bullets: ["Event A — leads / partners / intelligence collected.", "Event B — same."] },
    { heading: "Programme tie-ins",
      note: "Which Thoughtrons programmes (CHAITANYA, RAKESH, etc.) BD work supported this month, and what surfaced.",
      body: "Programme X: <BD activity, what it surfaced, programme team action>." },
  ],
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
