# Thoughtrons Word Add-in — V1.0

A Microsoft 365 Word add-in that adds a **Thoughtrons** ribbon tab with a snippet library, classification stamps, and quick-insert buttons.

```
Word ribbon
└── Thoughtrons
    ├── Snippet Library  ┊ [Big button] — opens a side pane with all 18 snippets
    ├── Quick Insert     ┊ Exec Summary · Action Items · Risk Table · Memo Header · References
    └── Classification   ┊ UNCLASSIFIED · INTERNAL · RESTRICTED · CONFIDENTIAL
```

**18 snippets across 5 categories:**

- **Headings** — Eyebrow label · Document title · H1 · H2
- **Content** — Executive Summary block · Context block · Key takeaway callout
- **Tables** — Action items · Decisions · Risk · Pricing · Memo header
- **Academic** — References block · Figure caption
- **Classification** — UNCLASSIFIED / INTERNAL / RESTRICTED / CONFIDENTIAL stamps (header & footer of every section)

## Architecture

Snippets are inserted programmatically via the Word JS API, not as base64 OOXML. This makes them small, fast, and trivial to extend:

- `src/snippets.js` — defines each snippet as a JS function that calls Word API `body.insertParagraph()` and `body.insertTable()` with consistent styling helpers.
- `src/commands.js` — wires 9 ribbon buttons (5 quick-insert + 4 classification) to snippet functions.
- `src/taskpane.html` — the side pane library, lists all 18 snippets with category tabs and search.
- `manifest.xml` — declares the ribbon, taskpane, and resources.

To add a new snippet: add it to `SNIPPETS` in `snippets.js`, then add it to the `LIBRARY` array in `taskpane.html`. It appears in the library next time anyone opens the pane — no manifest change needed.

## Deployment

Same path as PowerPoint: host the files on GitHub Pages (or any HTTPS host), find/replace the placeholder URL `https://thoughtrons.github.io/word-addin` to your real URL, then upload `manifest.xml` via Microsoft 365 admin centre → Integrated apps. See the unified deployment guide in `Brand_System_v1/README_DEPLOYMENT.md`.

## Files

```
word_addin/
├── manifest.xml
├── README.md
├── src/
│   ├── commands.html / commands.js
│   ├── snippets.js
│   └── taskpane.html
└── assets/icons/   ← 18 PNGs (3 sizes × 6 categories)
```
