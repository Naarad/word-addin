/*
 * Thoughtrons Word add-in — ribbon button handlers.
 *
 * The full library lives in the taskpane (taskpane.html). The ribbon only
 * exposes the most-used insert actions plus the 4 classification stamps.
 *
 * Adding a new snippet to the LIBRARY: just add it to snippets.js and to
 * the LIBRARY array in taskpane.html — no manifest change needed.
 *
 * Promoting a snippet to a RIBBON button: add a Button entry in manifest.xml
 * under GroupQuick, then add a matching Office.actions.associate(...) below.
 */

Office.onReady(function () {
  // Quick Insert (ribbon Group: Quick Insert)
  Office.actions.associate("insertExecSummary",        runner("exec_summary"));
  Office.actions.associate("insertActionItems",        runner("action_items_table"));
  Office.actions.associate("insertRiskTable",          runner("risk_table"));
  Office.actions.associate("insertMemoHeader",         runner("memo_header"));
  Office.actions.associate("insertReferences",         runner("references_block"));

  // Classification (ribbon Group: Classification > menu)
  Office.actions.associate("stampUnclassified",  runner("classify_unclassified"));
  Office.actions.associate("stampInternal",      runner("classify_internal"));
  Office.actions.associate("stampRestricted",    runner("classify_restricted"));
  Office.actions.associate("stampConfidential",  runner("classify_confidential"));
});

function runner(snippetKey) {
  return function (event) {
    const fn = (window.ThoughtronsSnippets || {})[snippetKey];
    if (!fn) {
      console.error("Thoughtrons add-in: unknown snippet", snippetKey);
      event.completed();
      return;
    }
    Word.run(ctx => Promise.resolve(fn(ctx)))
      .catch(err => console.error("Thoughtrons add-in error:", err))
      .finally(() => event.completed());
  };
}
